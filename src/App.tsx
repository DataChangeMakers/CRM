/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from "react";
import { GoogleGenAI } from "@google/genai";
import { 
  auth, signOut, onAuthStateChanged, db, 
  collection, query, onSnapshot, User, setDoc, doc, addDoc, deleteDoc, updateDoc, getDoc, orderBy, handleFirestoreError, OperationType, writeBatch,
  signInWithEmailAndPassword, createUserWithEmailAndPassword, updatePassword, sendPasswordResetEmail
} from "./firebase";
import { 
  researchNonProfit, generateFullAssessment, summarizeBrief, assessDataQuality, researchLinkedIn,
  extractActionsFromLeads, suggestFollowUpEmail, LeadAction,
  NonProfitResearch, FullAssessmentResult, DataQualityAssessment 
} from "./geminiService";
import { cn, formatDate } from "./lib/utils";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import { toPng, toJpeg } from "html-to-image";
import Papa from "papaparse";
import { extractTextFromDoc } from './lib/fileParser';
import { 
  LayoutDashboard, Search, FileText, CheckCircle2, AlertCircle, Plus, Check,
  Trash2, ExternalLink, ChevronRight, ChevronDown, BarChart3, Users, HeartHandshake, Loader2, LogOut, SearchIcon, Database,
  Target, DollarSign, Download, X, ArrowRightCircle, ArrowRight, Shield, ShieldCheck, ShieldAlert, Save, UploadCloud,
  Activity, MapPin, MessageSquare, User as UserIcon, Globe, AtSign, Calendar, Languages, CalendarDays, UserCircle, 
  Link, FastForward, Heart, Tag, Paperclip, GripVertical, ArrowUpRight, Edit3, Mail, Sparkles, Copy, RotateCw, Archive, RotateCcw,
  Cpu, Linkedin, Star, Fingerprint, Hash, Lock, Clock
} from "lucide-react";
import {
  DndContext,
  closestCenter,
  closestCorners,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  DragEndEvent,
  DragStartEvent,
  useDraggable,
  useDroppable,
  DragOverlay
} from '@dnd-kit/core';
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  verticalListSortingStrategy,
  useSortable
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { motion, AnimatePresence } from "framer-motion";
import { PieChart, Pie, Cell, ResponsiveContainer, RadarChart, PolarGrid, PolarAngleAxis, Radar } from "recharts";
import Markdown from 'react-markdown';

type View = "dashboard" | "research" | "assessments" | "briefs" | "leads" | "admin";

const parseRevenueValue = (rev: string): number => {
  if (!rev) return 0;
  const cleaned = rev.replace(/[^0-9.]/g, '');
  let val = parseFloat(cleaned);
  const lowRev = rev.toLowerCase();
  if (lowRev.includes('k')) val *= 1000;
  if (lowRev.includes('m')) val *= 1000000;
  if (lowRev.includes('b')) val *= 1000000000;
  return isNaN(val) ? 0 : val;
};

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [userProfile, setUserProfile] = useState<any>(null);
  const [isUnauthorized, setIsUnauthorized] = useState(false);
  const [settings, setSettings] = useState<any>(null);
  const [loading, setLoading] = useState(true);
  const [view, setView] = useState<View>("dashboard");
  const [users, setUsers] = useState<any[]>([]);

  const [impersonationRole, setImpersonationRole] = useState<string | null>(null);
  const [assessments, setAssessments] = useState<any[]>([]);
  const [showHelp, setShowHelp] = useState(false);
  const [researchLogs, setResearchLogs] = useState<any[]>([]);
  const [leads, setLeads] = useState<any[]>([]);
  const [selectedAssessmentId, setSelectedAssessmentId] = useState<string | null>(null);
  const [selectedResearchId, setSelectedResearchId] = useState<string | null>(null);
  const [notification, setNotification] = useState<{message: string, type: 'success' | 'error'} | null>(null);
  
  const notify = (message: string, type: 'success' | 'error' = 'success') => {
    setNotification({ message, type });
  };

  const logLeadEvent = async (leadId: string, type: string, description: string, metadata: any = {}) => {
    try {
      const uploader = userProfile?.displayName || userProfile?.name || auth.currentUser?.email || "System";
      const email = auth.currentUser?.email || "Unknown";
      const timestamp = new Date().toISOString();
      
      await addDoc(collection(db, "leads", leadId, "timeline"), {
        type,
        description,
        author: uploader,
        authorEmail: email,
        timestamp,
        metadata
      });
    } catch (err) {
      console.error("Failed to log lead event:", err);
    }
  };

  const [sessionApiKey, setSessionApiKey] = useState(() => localStorage.getItem('dcm_session_api_key') || '');
  const [sessionModel, setSessionModel] = useState(() => localStorage.getItem('dcm_session_model') || '');
  const [showSessionKeyModal, setShowSessionKeyModal] = useState(false);
  const [showChangePasswordModal, setShowChangePasswordModal] = useState(false);
  const [loginEmail, setLoginEmail] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [newPassword, setNewPassword] = useState("");
  const [confirmNewPassword, setConfirmNewPassword] = useState("");
  const [changePasswordError, setChangePasswordError] = useState("");
  const [isChangingPassword, setIsChangingPassword] = useState(false);
  
  const getAIConfig = () => {
    const key = sessionApiKey.trim() || settings?.aiResearch?.apiKey || "";
    const model = sessionModel || settings?.aiResearch?.model || "gemini-1.5-flash";
    return {
      apiKey: key,
      model: model,
      // Include a flag to indicate if we are using a personal session key
      isSessionKey: !!sessionApiKey.trim()
    };
  };

  const handleAiError = (err: any, source: string) => {
    console.error(`AI Error from ${source}:`, err);
    const isQuota = err.message?.includes("429") || err.message?.includes("quota") || err.message?.includes("RESOURCE_EXHAUSTED") || err.message?.includes("rate limit");
    
    if (isQuota) {
      const config = getAIConfig();
      if (config.isSessionKey) {
        setNotification({ 
          message: `Personal Key Quota: Your key is limited for ${config.model}. Try switching to 'Gemini 1.5 Flash' (balanced) in AI Config.`, 
          type: "error" 
        });
        // We only pop up if it's the first time in a while, or let the user decide
      } else {
        setNotification({ 
          message: "Global project quota reached. Please provide a personal API Key in 'AI Config'.", 
          type: "error" 
        });
        setShowSessionKeyModal(true);
      }
    } else {
      setNotification({ 
        message: err.message || `AI operation failed in ${source}.`, 
        type: "error" 
      });
    }
  };
  useEffect(() => {
    if (notification) {
      const timer = setTimeout(() => setNotification(null), 5000);
      return () => clearTimeout(timer);
    }
  }, [notification]);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (u) => {
      setUser(u);
      setLoading(false);
      if (!u) {
        setUsers([]);
        setAssessments([]);
        setResearchLogs([]);
        setLeads([]);
        setUserProfile(null);
        setSettings(null);
        setIsUnauthorized(false);
      }
    });
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    if (!user) return;

    const qUsers = query(collection(db, "appUsers"));
    const unsubUsers = onSnapshot(qUsers, (snap) => {
      const fetchedUsers = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      fetchedUsers.sort((a: any, b: any) => {
        const emailA = (a.email || "").toLowerCase();
        const emailB = (b.email || "").toLowerCase();
        return emailA.localeCompare(emailB);
      });
      setUsers(fetchedUsers);
    }, (err) => handleFirestoreError(err, OperationType.LIST, "appUsers"));

    const qAss = query(collection(db, "assessments"), orderBy("createdAt", "desc"));
    const unsubAss = onSnapshot(qAss, (snap) => {
      setAssessments(snap.docs.map(d => ({ id: d.id, ...d.data() })));
    }, (err) => handleFirestoreError(err, OperationType.LIST, "assessments"));

    const qLogs = query(collection(db, "researchLogs"), orderBy("timestamp", "desc"));
    const unsubLogs = onSnapshot(qLogs, (snap) => {
      setResearchLogs(snap.docs.map(d => ({ id: d.id, ...d.data() })));
    }, (err) => handleFirestoreError(err, OperationType.LIST, "researchLogs"));

    const qLeads = query(collection(db, "leads"), orderBy("organisation", "asc"));
    const unsubLeads = onSnapshot(qLeads, (snap) => {
      setLeads(snap.docs.map(d => ({ id: d.id, ...d.data() })));
    }, (err) => handleFirestoreError(err, OperationType.LIST, "leads"));

    // Listen to user profile
    const unsubProfile = onSnapshot(doc(db, "appUsers", user.uid), async (snap) => {
      if (snap.exists()) {
        const data = snap.data();
        setUserProfile({ id: snap.id, ...data });
        setIsUnauthorized(false);
      } else {
        // CHECK FOR INVITE/EMAIL-BASED PROFILE FIRST
        let invitedProfile: any = null;
        try {
          const emailLower = user.email?.toLowerCase() || "";
          const inviteRef = doc(db, "appUsers", emailLower);
          const inviteSnap = await getDoc(inviteRef);
          
          if (inviteSnap.exists()) {
            invitedProfile = inviteSnap.data();
            // Create the UID-indexed record using the invitation data
            await setDoc(doc(db, "appUsers", user.uid), {
              ...invitedProfile,
              uid: user.uid,
              displayName: invitedProfile.displayName || user.displayName || (emailLower || "").split('@')[0],
              email: user.email,
              updatedAt: new Date().toISOString()
            }, { merge: true });
            
            // Clean up the email-based record as it's no longer needed
            if (inviteRef.id !== user.uid) {
              await deleteDoc(inviteRef);
            }
          } else {
            // New user without invitation - CHECK FOR ADMIN DEFAULT
            const isAdminEmail = (user.email === "frederic@datachangemakers.org" || user.email === "fredericf.fery@gmail.com");
            
            if (isAdminEmail) {
              const newProfile = {
                email: user.email,
                displayName: user.displayName || user.email?.split('@')[0] || "Admin",
                firstName: "",
                lastName: "",
                role: "Admin",
                createdAt: new Date().toISOString()
              };
              await setDoc(doc(db, "appUsers", user.uid), newProfile);
              setUserProfile({ id: user.uid, ...newProfile });
              setIsUnauthorized(false);
            } else {
              // NOT AUTHORIZED
              setIsUnauthorized(true);
            }
          }
        } catch (e) {
          console.error("Profile synchronization error:", e);
        }
      }
    }, (err) => handleFirestoreError(err, OperationType.GET, `appUsers/${user.uid}`));

    // Listen to settings
    const unsubSettings = onSnapshot(doc(db, "settings", "global"), (snap) => {
      if (snap.exists()) {
        const data = snap.data() as any;
        // Auto-fix lead types if they are the old defaults
        const oldTypes = ["DCM Boards", "DCM Network", "Web search", "AI Search", "Other"];
        const currentTypes = data.leadsConfig?.types || [];
        if (currentTypes.length === oldTypes.length && currentTypes.every((v: string, i: number) => v === oldTypes[i])) {
          setDoc(doc(db, "settings", "global"), {
            leadsConfig: {
              ...data.leadsConfig,
              types: ["Education", "Partnership", "Project"]
            }
          }, { merge: true });
        }
        setSettings(data);
        
        // Auto-fix lead statuses if "Assessed" is missing
        const currentStatuses = data.leadsConfig?.statuses || [];
        if (currentStatuses.length > 0 && !currentStatuses.includes("Assessed")) {
           const newStatuses = [...currentStatuses];
           const index = newStatuses.indexOf("Under assessment");
           if (index !== -1) {
             newStatuses.splice(index + 1, 0, "Assessed");
           } else {
             newStatuses.push("Assessed");
           }
           setDoc(doc(db, "settings", "global"), {
             leadsConfig: {
               ...data.leadsConfig,
               statuses: newStatuses
             }
           }, { merge: true });
        }

        // Auto-correct invalid models to recommended ones
        // Auto-correct invalid models to recommended ones
        if (data.aiResearch?.model && data.aiResearch.model === "gemini-flash-latest") {
          setDoc(doc(db, "settings", "global"), {
            aiResearch: {
              ...data.aiResearch,
              model: "gemini-1.5-flash"
            }
          }, { merge: true });
        }
      } else {
        // Initialize default settings if missing
        const defaultSettings = {
          scoring: {
            verification: [
              { id: 'reputable', label: "Org is reputable Non-profit, charity or Social Enterprise (mission, website exist, CEO listed...)", type: 'boolean' },
              { id: 'fitsMission', label: "Org supports a cause that fits DCM mission [Data for Good]", type: 'boolean' },
              { id: 'isNonProfit', label: "Org is a Non-profit, charity or Social Enterprise (not a private company seeking profits)", type: 'boolean' },
              { id: 'available', label: "Org available to participate in a project in the next 2 months?", type: 'boolean', hint: "If one NO is answered, org is not a good fit for project or org is not ready (but we could revisit if conditions change)" }
            ],
            validation: [
              { id: 'dataQuality', label: "If org provided data or sample data, is it of acceptable/usable quality?", type: 'score' },
              { id: 'problemStatement', label: "Org has articulated clear problem statement and goals?", type: 'score' },
              { id: 'missionAlignment', label: "Org has articulated how they advance inclusion and diversity goals?", type: 'score' },
              { id: 'partnershipReason', label: "Org has articulated a good reason DCM should partner with them?", type: 'score' },
              { id: 'fundsAvailable', label: "Org has funds for a potential admin paid project", type: 'score' }
            ],
            validationChecks: [
              { id: 'diverseStaff', label: "Non-profit has a diverse staff", type: 'boolean' },
              { id: 'diverseDemographic', label: "Non-profit serves a diverse demographic", type: 'boolean' },
              { id: 'inclusiveMarketing', label: "Non-profit markets on a variety of platforms inclusive of various audiences", type: 'boolean' }
            ]
          },
          aiResearch: {
            sources: ["LinkedIn", "Google", "Website", "Reports"],
            model: "gemini-3-flash-preview"
          },
          leadsConfig: {
            statuses: ["Under assessment", "Assessed", "First meeting/contact", "Needs identifying", "Follow up", "Under consideration", "Approved for Future Project", "Completed", "Not Interest", "Not suitable"],
            types: ["Education", "Partnership", "Project"]
          },
          displayConfig: {
            dateFormat: "dd/mm/yyyy"
          }
        };
        setDoc(doc(db, "settings", "global"), defaultSettings);
        setSettings(defaultSettings);
      }
    }, (err) => handleFirestoreError(err, OperationType.GET, "settings/global"));

    return () => {
      unsubUsers();
      unsubAss();
      unsubLogs();
      unsubLeads();
      unsubProfile();
      unsubSettings();
    };
  }, [user, db]);

  const [selectedLeadIdForBrief, setSelectedLeadIdForBrief] = useState<string | null>(null);

  // Active profile is either the real one or the impersonated one
  const activeProfile = impersonationRole ? { ...userProfile, role: impersonationRole, isImpersonating: true } : userProfile;
  const isSystemAdmin = userProfile?.role?.toLowerCase() === "admin";
  const isAdmin = activeProfile?.role?.toLowerCase() === "admin";
  const isBoards = activeProfile?.role?.toLowerCase() === "dcm boards" || isAdmin;
  const isAssessor = activeProfile?.role?.toLowerCase() === "assessor" || isBoards;

  if (loading) {
    return (
      <div className="h-screen w-screen flex items-center justify-center bg-brand-bg">
        <Loader2 className="animate-spin text-brand-primary w-8 h-8" />
      </div>
    );
  }

  if (user && isUnauthorized) {
    return (
      <div className="h-screen w-screen flex flex-col items-center justify-center bg-brand-bg p-6 text-center">
        <motion.div 
          initial={{ opacity: 0, scale: 0.95 }}
          animate={{ opacity: 1, scale: 1 }}
          className="max-w-md bg-white p-12 rounded-2xl border border-brand-border shadow-xl text-center"
        >
          <div className="w-16 h-16 bg-red-50 text-red-600 rounded-full flex items-center justify-center mx-auto mb-6">
            <Lock className="w-8 h-8" />
          </div>
          <h1 className="text-2xl font-black text-brand-primary mb-4 tracking-tight uppercase">Access Denied</h1>
          <p className="text-brand-muted mb-8 leading-relaxed">
            Your account <span className="font-bold text-brand-primary">{user.email}</span> is not authorized to access this platform.
          </p>
          <div className="space-y-4">
            <p className="text-xs text-brand-muted font-medium italic">
              Please contact your administrator to request access.
            </p>
            <button 
              onClick={() => signOut(auth)}
              className="w-full bg-brand-primary text-white py-3 px-6 rounded-xl font-bold hover:bg-brand-primary/90 transition-all flex items-center justify-center gap-2"
            >
              Sign Out
            </button>
          </div>
        </motion.div>
      </div>
    );
  }

  const handleEmailLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoginError("");
    setIsLoggingIn(true);
    const emailLower = loginEmail.toLowerCase().trim();
    const password = loginPassword.trim();

    if (!emailLower || !password) {
      setLoginError("Please enter both email and password.");
      setIsLoggingIn(false);
      return;
    }

    try {
      // 1. First, try to sign in normally
      await signInWithEmailAndPassword(auth, emailLower, password);
      setIsLoggingIn(false);
    } catch (err: any) {
      console.error("Login attempt failed:", err);
      
      const isCredentialError = err.code === "auth/user-not-found" || 
                                err.code === "auth/invalid-credential" || 
                                err.code === "auth/wrong-password" ||
                                err.message?.includes("user-not-found") || 
                                err.message?.includes("invalid-credential") ||
                                err.message?.includes("wrong-password");

      if (isCredentialError) {
        // Check if the user is a default admin email or invited in appUsers
        const isAdminEmail = (emailLower === "frederic@datachangemakers.org" || 
                              emailLower === "fredericf.fery@gmail.com" || 
                              emailLower === "zifrench2@gmail.com");
        
        let shouldAutoRegister = isAdminEmail;
        let dbUserPassword = "Changemaker2026!";
        try {
          const inviteSnap = await getDoc(doc(db, "appUsers", emailLower));
          if (inviteSnap.exists()) {
            shouldAutoRegister = true;
            const data = inviteSnap.data();
            if (data && data.password) {
              dbUserPassword = data.password.trim();
            }
          }
        } catch (dbErr) {
          console.error("Failed to fetch invitation database record:", dbErr);
        }

        if (shouldAutoRegister) {
          // If the password used matches the configured password in database
          if (password === dbUserPassword) {
            try {
              await createUserWithEmailAndPassword(auth, emailLower, password);
              setIsLoggingIn(false);
              return;
            } catch (createErr: any) {
              if (createErr.code === "auth/email-already-in-use" || createErr.message?.includes("email-already-in-use")) {
                try {
                  await sendPasswordResetEmail(auth, emailLower);
                  setLoginError(`Your admin changed your password to '${password}' in our team directory. To match this on your active account, we have automatically sent an official Firebase password reset link to ${emailLower}. Please check your inbox (and spam folder) to set your official credential.`);
                } catch (resetErr) {
                  setLoginError("This email was already set up in Firebase. Since your administrator updated your password, please use the 'Forgot Password?' link below to reset your official login credential.");
                }
              } else if (createErr.code === "auth/operation-not-allowed" || createErr.message?.includes("operation-not-allowed")) {
                setLoginError("Email/Password provider is disabled in your Firebase console. Please go to the Firebase Console -> Authentication -> Sign-in method, click 'Add new provider', and enable 'Email/Password' to activate login functionality!");
              } else {
                setLoginError(createErr.message || "Account registration conflict. Please try again.");
              }
              setIsLoggingIn(false);
              return;
            }
          } else {
            setLoginError("Incorrect password. Please try again.");
            setIsLoggingIn(false);
            return;
          }
        } else {
          setLoginError("This email address is not pre-authorized to join this platform. Please contact your administrator to receive an invitation.");
          setIsLoggingIn(false);
          return;
        }
      }

      // Handle other common authentication failures
      if (err.code === "auth/invalid-email" || err.message?.includes("invalid-email")) {
        setLoginError("Please enter a valid email address.");
      } else if (err.code === "auth/operation-not-allowed" || err.message?.includes("operation-not-allowed")) {
        setLoginError("Email/Password provider is disabled in your Firebase console. Please go to the Firebase Console -> Authentication -> Sign-in method, click 'Add new provider', and enable 'Email/Password' to activate login functionality!");
      } else if (err.code === "auth/too-many-requests" || err.message?.includes("too-many-requests")) {
        setLoginError("Access to this account has been temporarily disabled due to many failed login attempts. You can immediately restore it by resetting your password or trying again later.");
      } else {
        setLoginError(err.message || "Failed to log in.");
      }
      setIsLoggingIn(false);
    }
  };

  if (!user) {
    return (
      <div className="h-screen w-screen flex flex-col items-center justify-center bg-brand-bg p-6 text-center">
        <motion.div 
          initial={{ opacity: 0, scale: 0.95 }}
          animate={{ opacity: 1, scale: 1 }}
          className="w-full max-w-md bg-white p-10 rounded-2xl border border-brand-border shadow-xl text-left"
        >
          <div className="flex items-center gap-3 mb-6 justify-center">
            <div className="w-10 h-10 bg-brand-accent rounded-xl flex items-center justify-center shadow-lg shadow-brand-accent/20">
              <span className="text-white text-base font-black">Δ</span>
            </div>
            <h1 className="text-2xl font-black tracking-tight text-brand-primary">Changemaker Agent</h1>
          </div>
          
          <p className="text-brand-muted text-sm text-center mb-8 leading-relaxed font-semibold">
            Empowering data changemakers to assess and validate high-impact non-profit projects.
          </p>

          <form onSubmit={handleEmailLogin} className="space-y-4">
            <div>
              <label className="block text-[10px] font-black uppercase tracking-widest text-brand-muted mb-1.5 pl-1">
                Email Address
              </label>
              <div className="relative">
                <AtSign className="absolute left-3.5 top-1/2 -translate-y-1/2 text-brand-muted/70 w-4 h-4" />
                <input 
                  type="email"
                  value={loginEmail}
                  onChange={(e) => setLoginEmail(e.target.value)}
                  placeholder="youremail@yourdomain.com"
                  className="w-full bg-brand-bg border border-brand-border rounded-xl py-3 pl-10 pr-4 text-xs font-bold text-brand-primary placeholder:text-brand-muted/40 focus:outline-none focus:border-brand-accent transition-colors"
                  required
                />
              </div>
            </div>

            <div>
              <div className="flex justify-between items-center mb-1.5 pl-1">
                <label className="block text-[10px] font-black uppercase tracking-widest text-brand-muted">
                  Password
                </label>
                <button 
                  type="button" 
                  onClick={async () => {
                    const emailLower = loginEmail.toLowerCase().trim();
                    if (!emailLower) {
                      setLoginError("Please enter your email address first in the input above, then click 'Forgot Password?' to receive a reset link.");
                      return;
                    }
                    try {
                      setLoginError("");
                      await sendPasswordResetEmail(auth, emailLower);
                      setLoginError(`A password reset link has been sent to ${emailLower}! Please check your email inbox and spam folder.`);
                    } catch (resetErr: any) {
                      if (resetErr.code === "auth/user-not-found" || resetErr.message?.includes("user-not-found")) {
                        setLoginError("This email is not pre-authorized or registered yet. Please contact your administrator.");
                      } else if (resetErr.code === "auth/invalid-email") {
                        setLoginError("Please enter a valid email address first.");
                      } else if (resetErr.code === "auth/operation-not-allowed" || resetErr.message?.includes("operation-not-allowed")) {
                        setLoginError("Email/Password provider is disabled in your Firebase console. Please enable it in Firebase Console -> Authentication -> Sign-in method.");
                      } else {
                        setLoginError(resetErr.message || "Failed to send reset email.");
                      }
                    }
                  }}
                  className="text-[10px] font-bold text-brand-accent hover:underline uppercase tracking-wide cursor-pointer"
                >
                  Forgot Password?
                </button>
              </div>
              <div className="relative">
                <Lock className="absolute left-3.5 top-1/2 -translate-y-1/2 text-brand-muted/70 w-4 h-4" />
                <input 
                  type="password"
                  value={loginPassword}
                  onChange={(e) => setLoginPassword(e.target.value)}
                  placeholder="••••••••••••"
                  className="w-full bg-brand-bg border border-brand-border rounded-xl py-3 pl-10 pr-4 text-xs font-bold text-brand-primary placeholder:text-brand-muted/40 focus:outline-none focus:border-brand-accent transition-colors"
                  required
                />
              </div>
            </div>

            {loginError && (
              <div className="p-3 bg-red-50 border border-red-100 rounded-xl flex gap-2 text-red-600 text-[11px] leading-relaxed font-bold">
                <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
                <span>{loginError}</span>
              </div>
            )}

            <button 
              type="submit"
              disabled={isLoggingIn}
              className="w-full bg-brand-primary text-white py-3.5 px-6 rounded-xl font-bold hover:bg-brand-primary/95 transition-all flex items-center justify-center gap-2 shadow-sm disabled:opacity-50 text-xs uppercase tracking-wider"
            >
              {isLoggingIn ? (
                <>
                  <Loader2 className="animate-spin w-4 h-4" /> Signing In...
                </>
              ) : "Sign In"}
            </button>
          </form>
        </motion.div>
      </div>
    );
  }

  const updateLeadsConfig = async (key: string, data: any) => {
    try {
      const finalData = Array.isArray(data) ? data.filter(Boolean) : data;
      await setDoc(doc(db, "settings", "global"), {
        leadsConfig: {
          ...(settings?.leadsConfig || {}),
          [key]: finalData
        }
      }, { merge: true });
    } catch (err) {
      handleFirestoreError(err, OperationType.UPDATE, "settings");
    }
  };

  const updateAIConfig = async (newConfig: any) => {
    try {
      await setDoc(doc(db, "settings", "global"), {
        aiResearch: newConfig
      }, { merge: true });
    } catch (err) {
      handleFirestoreError(err, OperationType.UPDATE, "settings");
    }
  };

  const updateScoringConfig = async (key: string, newConfig: any) => {
    try {
      await setDoc(doc(db, "settings", "global"), {
        scoring: {
          ...settings.scoring,
          [key]: newConfig
        }
      }, { merge: true });
    } catch (err) {
      handleFirestoreError(err, OperationType.UPDATE, "settings");
    }
  };

  const deletePartner = async (id: string) => {
    if (userProfile?.role !== "Admin" && userProfile?.role !== "DCM Boards") {
      setNotification({ message: "Only Admin or Board members can delete leads.", type: "error" });
      return;
    }
    
    if (!confirm("Are you sure you want to move this organization to the Archive? Related assessments will be permanently removed.")) return;

    try {
      const leadToArchive = leads.find(l => l.id === id);
      if (leadToArchive) {
        await setDoc(doc(db, "archivedLeads", id), {
          ...leadToArchive,
          archivedAt: new Date().toISOString(),
          archivedBy: userProfile?.email,
          archiveReason: "Manual deletion from dashboard"
        });
      }

      const batch = writeBatch(db);
      
      // Delete the lead doc
      batch.delete(doc(db, "leads", id));
      
      // Delete all related assessments
      const relatedAssessments = assessments.filter(a => a.nonProfitId === id);
      relatedAssessments.forEach(ass => {
        batch.delete(doc(db, "assessments", ass.id));
      });

      await batch.commit();
      setNotification({ message: "Lead moved to archive successfully.", type: "success" });
    } catch (err) {
      console.error("Delete partner error:", err);
      handleFirestoreError(err, OperationType.DELETE, "leads");
    }
  };

  return (
    <div className="flex h-screen bg-brand-bg text-brand-text font-sans selection:bg-brand-accent selection:text-white">
      <AnimatePresence>
        {notification && (
          <motion.div 
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -20 }}
            className={cn(
              "fixed top-6 left-1/2 -translate-x-1/2 z-[100] px-6 py-3 rounded-xl shadow-2xl flex items-center gap-3 border font-bold text-[11px] uppercase tracking-widest",
              notification.type === 'success' ? "bg-emerald-50 text-emerald-700 border-emerald-100" : "bg-red-50 text-red-700 border-red-100"
            )}
          >
            {notification.type === 'success' ? <CheckCircle2 className="w-4 h-4" /> : <AlertCircle className="w-4 h-4" />}
            {notification.message}
          </motion.div>
        )}
      </AnimatePresence>

      <HelpModal isOpen={showHelp} onClose={() => setShowHelp(false)} />
      
      <AnimatePresence>
        {showSessionKeyModal && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}
              onClick={() => setShowSessionKeyModal(false)}
              className="absolute inset-0 bg-brand-primary/40 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="bg-white rounded-3xl p-10 max-w-lg w-full relative z-10 card-shadow border border-brand-border"
            >
               <div className="flex justify-between items-start mb-6">
                 <div>
                   <h3 className="text-2xl font-black text-brand-primary tracking-tight uppercase">Session Intelligence</h3>
                   <p className="text-[10px] font-bold text-brand-muted uppercase tracking-widest mt-1">Override global API configuration for this browser</p>
                 </div>
                 <button onClick={() => setShowSessionKeyModal(false)} className="p-2 hover:bg-brand-bg rounded-full transition-colors">
                   <X className="w-5 h-5 text-brand-muted" />
                 </button>
               </div>

               <div className="space-y-6">
                  <div className="p-4 bg-brand-bg rounded-2xl border border-brand-border">
                    <div className="flex items-center gap-2 mb-2">
                       <ShieldCheck className="w-4 h-4 text-emerald-500" />
                       <span className="text-[10px] font-black uppercase text-brand-primary">Config Status</span>
                    </div>
                    <p className="text-[11px] font-medium text-brand-primary leading-relaxed">
                      {sessionApiKey ? (
                        <>You are currently using a <span className="text-emerald-600 font-black">Personal Session Key</span>. This overrides the system-wide configuration for your current browser.</>
                      ) : (
                        <>You are using the <span className="text-brand-accent font-black">Global Project Key</span>. If you hit quota limits, providing a personal key will allow you to continue.</>
                      )}
                    </p>
                  </div>

                  <div className="space-y-4">
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-brand-primary uppercase tracking-widest">Personal Gemini API Key</label>
                      <div className="relative">
                        <input 
                          type="text"
                          placeholder={sessionApiKey ? "Key currently active..." : "Paste your personal key (AI Studio API Key)..."}
                          value={sessionApiKey}
                          onChange={(e) => {
                            const val = e.target.value.trim();
                            setSessionApiKey(val);
                            if (val) {
                              localStorage.setItem('dcm_session_api_key', val);
                            }
                          }}
                          className="w-full bg-white border border-brand-border rounded-xl p-4 pr-12 text-sm font-bold outline-none focus:ring-1 focus:ring-brand-accent shadow-sm"
                        />
                        <Sparkles className="absolute right-4 top-1/2 -translate-y-1/2 w-4 h-4 text-brand-accent" />
                      </div>
                      {sessionApiKey && (
                        <div className="flex items-center justify-between">
                          <p className="text-[9px] font-bold text-emerald-600 uppercase flex items-center gap-1">
                            <Check className="w-3 h-3" /> Key Stored Locally
                          </p>
                          <button 
                            onClick={async () => {
                              try {
                                const ai = new GoogleGenAI({ apiKey: sessionApiKey });
                                const result = await ai.models.generateContent({
                                  model: "gemini-3-flash-preview",
                                  contents: "ping",
                                });
                                if (result.text) {
                                  setNotification({ message: "API Key Verified - Ready for use", type: "success" });
                                } else {
                                  setNotification({ message: "Key Test Failed: Invalid response", type: "error" });
                                }
                              } catch (err: any) {
                                setNotification({ message: `Key Test Failed: ${err.message || "Invalid Key"}`, type: "error" });
                              }
                            }}
                            className="text-[9px] font-black text-brand-accent uppercase hover:underline"
                          >
                            Test Connection
                          </button>
                        </div>
                      )}
                    </div>

                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-brand-primary uppercase tracking-widest">Override Model</label>
                      <select 
                        value={sessionModel || (settings?.aiResearch?.model || "gemini-1.5-flash")}
                        onChange={(e) => {
                          const val = e.target.value;
                          setSessionModel(val);
                          localStorage.setItem('dcm_session_model', val);
                        }}
                        className="w-full bg-white border border-brand-border rounded-xl p-4 text-xs font-bold outline-none focus:ring-1 focus:ring-brand-accent shadow-sm appearance-none cursor-pointer"
                      >
                         <optgroup label="Stable Models (Recommended)">
                            <option value="gemini-1.5-flash">Gemini 1.5 Flash (Fast)</option>
                            <option value="gemini-1.5-pro">Gemini 1.5 Pro (Powerful)</option>
                         </optgroup>
                         <optgroup label="Advanced (Intelligence Models)">
                            <option value="gemini-3-flash-preview">Gemini 3 Flash (Latest)</option>
                            <option value="gemini-3.1-pro-preview">Gemini 3.1 Pro (Heavy Tasks)</option>
                         </optgroup>
                      </select>
                      <p className="text-[9px] font-bold text-brand-muted uppercase">Recommended: 'Gemini 1.5 Flash' for optimal performance and stability.</p>
                    </div>
                  </div>

                  <div className="flex gap-4 pt-2">
                    <button 
                      onClick={() => {
                        setSessionApiKey("");
                        setSessionModel("");
                        localStorage.removeItem('dcm_session_api_key');
                        localStorage.removeItem('dcm_session_model');
                        setNotification({ message: "Switched back to global configuration", type: "success" });
                        setShowSessionKeyModal(false);
                      }}
                      className="flex-1 px-6 py-4 bg-white border border-brand-border text-brand-muted text-[11px] font-black uppercase tracking-widest rounded-xl hover:bg-brand-bg transition-all"
                    >
                      Clear & Use Global
                    </button>
                    <button 
                      onClick={() => {
                        setNotification({ message: "Session key activated", type: "success" });
                        setShowSessionKeyModal(false);
                      }}
                      className="flex-1 px-6 py-4 bg-brand-accent text-white text-[11px] font-black uppercase tracking-widest rounded-xl hover:bg-brand-accent/90 transition-all shadow-lg shadow-brand-accent/20"
                    >
                      Save Configuration
                    </button>
                  </div>
               </div>
            </motion.div>
          </div>
        )}

        {showChangePasswordModal && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}
              onClick={() => {
                setNewPassword("");
                setConfirmNewPassword("");
                setChangePasswordError("");
                setShowChangePasswordModal(false);
              }}
              className="absolute inset-0 bg-brand-primary/40 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="bg-white rounded-3xl p-10 max-w-md w-full relative z-10 card-shadow border border-brand-border text-left"
            >
               <div className="flex justify-between items-start mb-6">
                 <div>
                   <h3 className="text-2xl font-black text-brand-primary tracking-tight uppercase">Change Password</h3>
                   <p className="text-[10px] font-bold text-brand-muted uppercase tracking-widest mt-1">Update your security credentials</p>
                 </div>
                 <button 
                  onClick={() => {
                    setNewPassword("");
                    setConfirmNewPassword("");
                    setChangePasswordError("");
                    setShowChangePasswordModal(false);
                  }} 
                  className="p-2 hover:bg-brand-bg rounded-full transition-colors"
                 >
                   <X className="w-5 h-5 text-brand-muted" />
                 </button>
               </div>

               <form onSubmit={async (e) => {
                 e.preventDefault();
                 setChangePasswordError("");
                 setIsChangingPassword(true);

                 if (newPassword.length < 6) {
                   setChangePasswordError("Password must be at least 6 characters long.");
                   setIsChangingPassword(false);
                   return;
                 }

                 if (newPassword !== confirmNewPassword) {
                   setChangePasswordError("New passwords do not match.");
                   setIsChangingPassword(false);
                   return;
                 }

                 try {
                   if (auth.currentUser) {
                     await updatePassword(auth.currentUser, newPassword);
                     setNotification({ message: "Password updated successfully!", type: "success" });
                     setNewPassword("");
                     setConfirmNewPassword("");
                     setShowChangePasswordModal(false);
                   } else {
                     setChangePasswordError("No authenticated user session found.");
                   }
                 } catch (err: any) {
                   console.error("Change password failed:", err);
                   if (err.message?.includes("requires-recent-login") || err.code === "auth/requires-recent-login") {
                     setChangePasswordError("For security, changing passwords requires recent login. Please log out and sign back in to complete this action.");
                   } else {
                     setChangePasswordError(err.message || "Failed to update password.");
                   }
                 }
                 setIsChangingPassword(false);
               }} className="space-y-4">
                  <div>
                    <label className="block text-[10px] font-black uppercase tracking-widest text-brand-muted mb-1.5 pl-1">
                      New Password
                    </label>
                    <div className="relative">
                      <Lock className="absolute left-3.5 top-1/2 -translate-y-1/2 text-brand-muted/70 w-4 h-4" />
                      <input 
                        type="password"
                        value={newPassword}
                        onChange={(e) => setNewPassword(e.target.value)}
                        placeholder="At least 6 characters"
                        className="w-full bg-brand-bg border border-brand-border rounded-xl py-3 pl-10 pr-4 text-xs font-bold text-brand-primary placeholder:text-brand-muted/40 focus:outline-none focus:border-brand-accent transition-colors"
                        required
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-[10px] font-black uppercase tracking-widest text-brand-muted mb-1.5 pl-1">
                      Confirm New Password
                    </label>
                    <div className="relative">
                      <Lock className="absolute left-3.5 top-1/2 -translate-y-1/2 text-brand-muted/70 w-4 h-4" />
                      <input 
                        type="password"
                        value={confirmNewPassword}
                        onChange={(e) => setConfirmNewPassword(e.target.value)}
                        placeholder="Repeat new password"
                        className="w-full bg-brand-bg border border-brand-border rounded-xl py-3 pl-10 pr-4 text-xs font-bold text-brand-primary placeholder:text-brand-muted/40 focus:outline-none focus:border-brand-accent transition-colors"
                        required
                      />
                    </div>
                  </div>

                  {changePasswordError && (
                    <div className="p-3 bg-red-50 border border-red-100 rounded-xl flex gap-2 text-red-600 text-[11px] leading-relaxed font-bold">
                      <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
                      <span>{changePasswordError}</span>
                    </div>
                  )}

                  <div className="flex gap-3 pt-2">
                    <button 
                      type="button"
                      onClick={() => {
                        setNewPassword("");
                        setConfirmNewPassword("");
                        setChangePasswordError("");
                        setShowChangePasswordModal(false);
                      }}
                      className="flex-1 px-4 py-3 bg-white border border-brand-border text-brand-muted text-[10px] font-black uppercase tracking-widest rounded-xl hover:bg-brand-bg transition-all"
                    >
                      Cancel
                    </button>
                    <button 
                      type="submit"
                      disabled={isChangingPassword}
                      className="flex-1 px-4 py-3 bg-brand-accent text-white text-[10px] font-black uppercase tracking-widest rounded-xl hover:bg-brand-accent/90 transition-all shadow-lg shadow-brand-accent/20 flex items-center justify-center gap-2 font-bold"
                    >
                      {isChangingPassword ? (
                        <>
                          <Loader2 className="animate-spin w-3.5 h-3.5" /> Updating...
                        </>
                      ) : "Update Password"}
                    </button>
                  </div>
               </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Sidebar */}
      <nav className="w-60 bg-brand-primary text-white flex flex-col pt-10 px-6 relative">
        <div className="mb-12">
          <div className="flex items-center gap-3 mb-2">
            <div className="w-9 h-9 bg-brand-accent rounded-lg flex items-center justify-center shadow-lg shadow-brand-accent/20">
              <span className="text-white text-sm font-black">Δ</span>
            </div>
            <span className="font-extrabold tracking-tight text-xl">Data ChangeMakers</span>
          </div>
          <p className="text-[10px] font-bold text-white/40 uppercase tracking-widest pl-12 line-clamp-1">Assessment Workspace</p>
        </div>

        <div className="flex-1 space-y-2">
          <NavItem active={view === "leads"} onClick={() => { setView("leads"); setSelectedResearchId(null); }} icon={HeartHandshake} label="DCM Leads" />
          <NavItem active={view === "dashboard"} onClick={() => { setView("dashboard"); setSelectedResearchId(null); }} icon={LayoutDashboard} label="Workplace" />
          <NavItem active={view === "research"} onClick={() => setView("research")} icon={Search} label="AI Research" />
          <NavItem active={view === "assessments"} onClick={() => { setView("assessments"); setSelectedResearchId(null); }} icon={CheckCircle2} label="Scoring" />
          <NavItem active={view === "briefs"} onClick={() => { setView("briefs"); setSelectedResearchId(null); }} icon={FileText} label="Assessment Pack" />
          {isAdmin && (
            <NavItem active={view === "admin"} onClick={() => setView("admin")} icon={Shield} label="Admin Area" />
          )}
          
          <button 
            onClick={() => setShowHelp(true)}
            className="w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all group text-white/60 hover:text-white hover:bg-white/5"
            title="Assessment Process Documentation"
          >
            <AlertCircle className="w-4 h-4" />
            <span className="text-[11px] font-black uppercase tracking-widest text-left">Help</span>
          </button>

          <button 
            onClick={() => setShowSessionKeyModal(true)}
            className="w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all group text-white/60 hover:text-white hover:bg-white/5"
            title="Session AI Configuration"
          >
            <Sparkles className="w-4 h-4 text-brand-accent/60 group-hover:text-brand-accent" />
            <span className="text-[11px] font-black uppercase tracking-widest text-left">AI Config</span>
            {sessionApiKey && <div className="w-2 h-2 rounded-full bg-emerald-500 ml-auto shadow-sm" />}
          </button>
        </div>

        <div className="py-8 border-t border-white/10 mt-auto">
          {isSystemAdmin && (
            <div className="mb-4 p-3 bg-white/5 rounded-lg border border-white/10">
              <p className="text-[9px] font-black uppercase tracking-widest text-white/40 mb-2">Impersonation Tool</p>
              <div className="flex flex-col gap-1">
                <button 
                  onClick={() => setImpersonationRole(null)}
                  className={cn("text-[9px] font-bold text-left px-2 py-1 rounded transition-all", impersonationRole === null ? "bg-white/20 text-white" : "text-white/40 hover:text-white")}
                >
                  Real Role (Admin)
                </button>
                <button 
                  onClick={() => setImpersonationRole("Assessor")}
                  className={cn("text-[9px] font-bold text-left px-2 py-1 rounded transition-all", impersonationRole === "Assessor" ? "bg-white/20 text-white" : "text-white/40 hover:text-white")}
                >
                  Impersonate Assessor
                </button>
                <button 
                  onClick={() => setImpersonationRole("DCM Boards")}
                  className={cn("text-[9px] font-bold text-left px-2 py-1 rounded transition-all", impersonationRole === "DCM Boards" ? "bg-white/20 text-white" : "text-white/40 hover:text-white")}
                >
                  Impersonate Boards
                </button>
              </div>
            </div>
          )}

          <div className="flex items-center gap-3 mb-4">
            {user.photoURL ? (
              <img src={user.photoURL} alt={user.displayName || ""} className="w-9 h-9 rounded-xl border border-white/20" />
            ) : (
              <div className="w-9 h-9 bg-brand-accent/20 rounded-xl border border-white/20 flex items-center justify-center shrink-0">
                <span className="text-brand-accent text-sm font-black">
                  {(user.displayName || user.email || "U").substring(0, 1).toUpperCase()}
                </span>
              </div>
            )}
            <div className="overflow-hidden">
              <div className="flex items-center gap-2">
                <p className="text-xs font-bold truncate">{user.displayName || user.email?.split('@')[0]}</p>
                {userProfile?.role && (
                  <span className="px-1.5 py-0.5 rounded bg-brand-accent/20 text-brand-accent text-[8px] font-black uppercase tracking-widest">
                    {userProfile.role}
                  </span>
                )}
              </div>
              <p className="text-[10px] text-white/40 truncate font-mono">{user.email}</p>
            </div>
          </div>
          <div className="space-y-2">
            <button 
              onClick={() => setShowChangePasswordModal(true)}
              className="w-full flex items-center gap-2 text-[10px] font-bold uppercase tracking-widest text-white/40 hover:text-brand-accent transition-colors"
            >
              <Lock className="w-3 h-3" /> Change Password
            </button>
            <button 
              onClick={() => signOut(auth)}
              className="w-full flex items-center gap-2 text-[10px] font-bold uppercase tracking-widest text-white/40 hover:text-red-400 transition-colors"
            >
              <LogOut className="w-3 h-3" /> Sign Out
            </button>
          </div>
        </div>
      </nav>

      {/* Main Content */}
      <main className="flex-1 overflow-y-auto overflow-x-auto p-10 scroll-smooth">
        <div className="mb-10 flex justify-between items-start">
          <div>
            <h2 className="text-3xl font-extrabold tracking-tight text-brand-primary mb-1">
              {view === "dashboard" && "Assessment Workspace"}
              {view === "leads" && "DCM Leads CRM"}
              {view === "research" && "Organization Discovery"}
              {view === "assessments" && "Scoring Dashboard"}
              {view === "briefs" && "Assessment Pack"}
            </h2>
            <p className="text-sm font-medium text-brand-muted uppercase tracking-wide">
              {view === "dashboard" && "Evaluating readiness for Data-for-Good projects"}
              {view === "leads" && "Managing partner relationships and pipeline"}
              {view === "research" && "Uncovering revenue and impact indicators"}
              {view === "assessments" && "Assessment Template Scoring"}
              {view === "briefs" && "Accessing generated recommendation and onboarding requirements"}
            </p>
          </div>
          {view === "research" && (
            <div className="w-80 relative group">
              <SearchIcon className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-brand-muted" />
              <input 
                type="text" 
                placeholder="Search organizations..." 
                className="w-full bg-white border border-brand-border rounded-lg py-2 pl-9 pr-4 text-sm font-medium focus:ring-2 focus:ring-brand-accent/20 outline-none transition-all"
              />
            </div>
          )}
        </div>

        <AnimatePresence mode="wait">
          {view === "dashboard" && (
            <Dashboard 
              assessments={assessments} 
              leads={leads}
              onSelectAssessment={(id) => {
                setSelectedAssessmentId(id);
                setView("assessments");
              }}
              onSelectResearch={(id) => {
                setSelectedResearchId(id);
                setView("research");
              }}
              onDeletePartner={deletePartner}
              userProfile={activeProfile}
            />
          )}
          {view === "leads" && (
            <LeadsView 
              leads={leads} 
              assessments={assessments}
              settings={settings} 
              userProfile={activeProfile} 
              aiConfig={getAIConfig()}
              onShowAiConfig={() => setShowSessionKeyModal(true)}
              handleAiError={handleAiError}
              notify={(msg, type) => setNotification({ message: msg, type })}
              updateLeadsConfig={updateLeadsConfig}
              logLeadEvent={logLeadEvent}
              onSelectResearch={(id) => {
                setSelectedResearchId(id);
                setView("research");
              }}
              users={users}
            />
          )}
          {view === "research" && (
            <ResearchTool 
              leads={leads}
              userProfile={activeProfile}
              initialPartnerId={selectedResearchId}
              onPartnerSelect={setSelectedResearchId}
              onDeletePartner={deletePartner}
              settings={settings}
              aiConfig={getAIConfig()}
              onShowAiConfig={() => setShowSessionKeyModal(true)}
              handleAiError={handleAiError}
              notify={(msg, type) => setNotification({ message: msg, type })}
              logLeadEvent={logLeadEvent}
            />
          )}
          {view === "assessments" && (
            <AssessmentsView 
              assessments={assessments} 
              leads={leads}
              selectedId={selectedAssessmentId}
              onClearSelection={() => setSelectedAssessmentId(null)}
              settings={settings}
              userProfile={activeProfile}
              aiConfig={getAIConfig()}
              onShowAiConfig={() => setShowSessionKeyModal(true)}
              handleAiError={handleAiError}
              notify={(msg, type) => setNotification({ message: msg, type })}
              onSetView={(v, leadId) => {
                setView(v);
                if (leadId) setSelectedLeadIdForBrief(leadId);
              }}
            />
          )}
          {view === "briefs" && (
            <BriefsView 
              assessments={assessments}
              leads={leads}
              settings={settings}
              userProfile={activeProfile}
              aiConfig={getAIConfig()}
              onShowAiConfig={() => setShowSessionKeyModal(true)}
              handleAiError={handleAiError}
              notify={(msg, type) => setNotification({ message: msg, type })}
              initialLeadId={selectedLeadIdForBrief}
              onClearSelection={() => setSelectedLeadIdForBrief(null)}
              logLeadEvent={logLeadEvent}
            />
          )}
          {view === "admin" && isAdmin && (
            <AdminView 
              leads={leads} 
              assessments={assessments}
              settings={settings} 
              userProfile={activeProfile} 
              updateLeadsConfig={updateLeadsConfig}
              updateAIConfig={updateAIConfig}
              updateScoringConfig={updateScoringConfig}
              notify={(msg, type) => setNotification({ message: msg, type })}
              sessionApiKey={sessionApiKey}
              users={users}
            />
          )}
        </AnimatePresence>
      </main>
    </div>
  );
}

function AdminView({ leads, assessments, settings, userProfile, updateLeadsConfig, updateAIConfig, updateScoringConfig, notify, sessionApiKey, users }: { 
  leads: any[], 
  assessments: any[],
  settings: any, 
  userProfile: any, 
  updateLeadsConfig: (key: string, data: any) => Promise<void>,
  updateAIConfig: (newConfig: any) => Promise<void>,
  updateScoringConfig: (key: string, newConfig: any) => Promise<void>,
  notify: (msg: string, type: 'success' | 'error') => void,
  sessionApiKey: string,
  users: any[]
}) {
  const [userSearchTerm, setUserSearchTerm] = useState("");
  const [activeAdminTab, setActiveAdminTab] = useState<"users" | "scoring" | "ai" | "leads" | "data" | "archive" | "permissions">("users");
  const [archivedLeads, setArchivedLeads] = useState<any[]>([]);
  const [newUserEmail, setNewUserEmail] = useState("");
  const [newUserFirstName, setNewUserFirstName] = useState("");
  const [newUserLastName, setNewUserLastName] = useState("");
  const [newUserPassword, setNewUserPassword] = useState("");
  const [newUserRole, setNewUserRole] = useState("DCM Boards");
  const [sendingEmail, setSendingEmail] = useState(false);
  const [invitedUser, setInvitedUser] = useState<any | null>(null);
  const [csvText, setCsvText] = useState("");
  const [isImporting, setIsImporting] = useState(false);
  const [loading, setLoading] = useState(false);
  const [confirmPurgeId, setConfirmPurgeId] = useState<string | null>(null);
  const [confirmGlobalPurge, setConfirmGlobalPurge] = useState(false);
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);

  const [newSource, setNewSource] = useState("");
  const [apiStatus, setApiStatus] = useState<any>(null);
  const [isCheckInProgress, setIsCheckInProgress] = useState(false);
  const [newStatus, setNewStatus] = useState("");
  const [newLeadSource, setNewLeadSource] = useState("");
  const [newLeadType, setNewLeadType] = useState("");

  const updateDisplayConfig = async (newConfig: any) => {
    try {
      await setDoc(doc(db, "settings", "global"), {
        displayConfig: newConfig
      }, { merge: true });
    } catch (err) {
      handleFirestoreError(err, OperationType.UPDATE, "settings");
    }
  };
  
  // User Editing State
  const [isEditingUser, setIsEditingUser] = useState(false);
  const [editingUser, setEditingUser] = useState<any | null>(null);

  useEffect(() => {
    if (activeAdminTab === "archive") {
      const q = query(collection(db, "archivedLeads"), orderBy("archivedAt", "desc"));
      const unsub = onSnapshot(q, (snap) => {
        setArchivedLeads(snap.docs.map(d => ({ id: d.id, ...d.data() })));
      }, (err) => handleFirestoreError(err, OperationType.LIST, "archivedLeads"));
      return () => unsub();
    }
  }, [activeAdminTab]);

  const addUser = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newUserEmail.trim()) return;
    
    const emailLower = newUserEmail.toLowerCase().trim();
    
    // Duplicate check
    if (users.some(u => u.email?.toLowerCase() === emailLower)) {
      notify("User with this email already exists", "error");
      return;
    }

    try {
      setSendingEmail(true);
      const userData = {
        email: emailLower,
        role: newUserRole,
        firstName: newUserFirstName.trim(),
        lastName: newUserLastName.trim(),
        displayName: newUserFirstName ? `${newUserFirstName.trim()} ${newUserLastName.trim()}`.trim() : (emailLower || "").split('@')[0],
        password: newUserPassword.trim() || "Changemaker2026!",
        createdAt: new Date().toISOString()
      };
      await setDoc(doc(db, "appUsers", emailLower), userData, { merge: true });

      setInvitedUser(userData);
      notify("User invited successfully", "success");
      
      setNewUserEmail("");
      setNewUserFirstName("");
      setNewUserLastName("");
      setNewUserPassword("");
      setSendingEmail(false);
    } catch (err) {
      notify("Failed to invite user", "error");
      handleFirestoreError(err, OperationType.CREATE, "appUsers");
      setSendingEmail(false);
    }
  };

  const deleteUser = async (userId: string) => {
    if (confirmDeleteId !== userId) {
      setConfirmDeleteId(userId);
      setTimeout(() => setConfirmDeleteId(null), 3000);
      return;
    }

    try {
      const userToDelete = users.find(u => u.id === userId);
      await deleteDoc(doc(db, "appUsers", userId));
      
      // If we deleted a UID record, also try to delete the invitation record for that email
      if (userToDelete?.email && userToDelete.email !== userId) {
        try {
          await deleteDoc(doc(db, "appUsers", userToDelete.email.toLowerCase().trim()));
        } catch (e) {
          // Ignore errors if invitation doesn't exist
        }
      }
      
      notify("User access removed", "success");
      setConfirmDeleteId(null);
    } catch (err) {
      notify("Failed to remove user", "error");
      handleFirestoreError(err, OperationType.DELETE, "appUsers");
    }
  };

  const updateUserRole = async (userId: string, role: string) => {
    try {
      await setDoc(doc(db, "appUsers", userId), { role }, { merge: true });
      notify("User role successfully updated", "success");
    } catch (err) {
      notify("Failed to update role", "error");
      handleFirestoreError(err, OperationType.UPDATE, "appUsers");
    }
  };

  const [selectedPurgeLeadId, setSelectedPurgeLeadId] = useState("");

  const purgeLeadResearch = async (leadId: string) => {
    if (!leadId) return;
    if (confirmPurgeId !== leadId) {
      setConfirmPurgeId(leadId);
      setTimeout(() => setConfirmPurgeId(null), 3000);
      return;
    }
    
    console.log("Purging lead research for:", leadId);
    const lead = leads.find(l => l.id === leadId);
    if (!lead) {
      console.warn("No lead found for ID:", leadId);
      return;
    }
    
    setLoading(true);
    try {
      await updateDoc(doc(db, "leads", leadId), {
        briefSummary: "",
        revenue: "",
        ein: "",
        charity_navigator_rating: "",
        propublica_grants: "",
        linkedin_overview: "",
        staff_linkedin_summary: "",
        staff_members: "",
        dcmComment: "",
        updatedAt: new Date().toISOString()
      });
      notify(`Research purged for ${lead.organisation}`, "success");
      setSelectedPurgeLeadId("");
      setConfirmPurgeId(null);
    } catch (err) {
      console.error("Purge failed:", err);
      notify("Wipe failed", "error");
    } finally {
      setLoading(false);
    }
  };

  const purgeAllAIResearch = async () => {
    if (!confirmGlobalPurge) {
      setConfirmGlobalPurge(true);
      setTimeout(() => setConfirmGlobalPurge(false), 3000);
      return;
    }
    
    console.log("Purging all research...");
    if (leads.length === 0) {
      notify("No leads to purge", "error");
      return;
    }
    
    setLoading(true);
    try {
      const batch = writeBatch(db);
      leads.forEach(l => {
        batch.update(doc(db, "leads", l.id), {
          briefSummary: "",
          revenue: "",
          ein: "",
          charity_navigator_rating: "",
          propublica_grants: "",
          linkedin_overview: "",
          staff_linkedin_summary: "",
          staff_members: "",
          dcmComment: "",
          updatedAt: new Date().toISOString()
        });
      });
      await batch.commit();
      notify("Global research database sanitized successfully.", "success");
      setConfirmGlobalPurge(false);
    } catch (err) {
      console.error("Global purge failed:", err);
      notify("Bulk purge failed", "error");
    } finally {
      setLoading(false);
    }
  };

  const [duplicateGroups, setDuplicateGroups] = useState<any[]>([]);

  const sensors = useSensors(
    useSensor(PointerSensor),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  const findDuplicates = () => {
    console.log("Scanning for duplicates among", leads.length, "leads");
    const groups: { [name: string]: any[] } = {};
    leads.forEach(l => {
      const name = (l.organisation || "").trim().toLowerCase();
      if (!groups[name]) groups[name] = [];
      groups[name].push(l);
    });
    
    const dups = Object.values(groups).filter(g => g.length > 1);
    setDuplicateGroups(dups);
    if (dups.length === 0) notify("No duplicate organisation names found.", "success");
    else notify(`Found ${dups.length} groups of duplicates.`, "success");
  };

  const deleteDuplicateLead = async (leadId: string, orgName: string) => {
    if (confirmDeleteId !== leadId) {
      setConfirmDeleteId(leadId);
      setTimeout(() => setConfirmDeleteId(null), 3000);
      return;
    }
    
    console.log("Archiving duplicate lead:", leadId, orgName);
    setLoading(true);
    try {
      const leadToArchive = leads.find(l => l.id === leadId);
      if (leadToArchive) {
        await setDoc(doc(db, "archivedLeads", leadId), {
          ...leadToArchive,
          archivedAt: new Date().toISOString(),
          archivedBy: userProfile?.email,
          archiveReason: "Duplicate resolution"
        });
      }
      
      const batch = writeBatch(db);
      batch.delete(doc(db, "leads", leadId));
      
      // Clean up orphaned assessments
      assessments.filter(a => a.nonProfitId === leadId).forEach(ass => {
        batch.delete(doc(db, "assessments", ass.id));
      });
      
      await batch.commit();
      
      setDuplicateGroups(prev => 
        prev.map(g => g.filter((l: any) => l.id !== leadId)).filter(g => g.length > 1)
      );
      notify("Entry moved to archive.", "success");
      setConfirmDeleteId(null);
    } catch (err) {
      console.error("Duplicate archiving failed:", err);
      notify("Archiving failed", "error");
    } finally {
      setLoading(false);
    }
  };

  const restoreLead = async (archivedLead: any) => {
    setLoading(true);
    try {
      const { id, archivedAt, archivedBy, archiveReason, ...leadData } = archivedLead;
      await setDoc(doc(db, "leads", id), {
        ...leadData,
        updatedAt: new Date().toISOString()
      });
      await deleteDoc(doc(db, "archivedLeads", id));
      notify("Lead restored successfully", "success");
    } catch (err) {
      notify("Failed to restore lead", "error");
      handleFirestoreError(err, OperationType.WRITE, "leads");
    } finally {
      setLoading(false);
    }
  };

  const permanentlyDeleteLead = async (id: string) => {
    if (confirmDeleteId !== id) {
      setConfirmDeleteId(id);
      setTimeout(() => setConfirmDeleteId(null), 3000);
      return;
    }
    setLoading(true);
    try {
      await deleteDoc(doc(db, "archivedLeads", id));
      notify("Lead permanently deleted", "success");
      setConfirmDeleteId(null);
    } catch (err) {
      notify("Deletion failed", "error");
    } finally {
      setLoading(false);
    }
  };

  const handleUpdateUser = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingUser) return;
    try {
      const updateData: any = {
        firstName: editingUser.firstName || "",
        lastName: editingUser.lastName || "",
        email: editingUser.email || "",
        role: editingUser.role || "Assessor"
      };
      if (editingUser.password !== undefined) {
        updateData.password = editingUser.password;
      }
      await setDoc(doc(db, "appUsers", editingUser.id), updateData, { merge: true });
      notify("User profile updated", "success");
      setIsEditingUser(false);
      setEditingUser(null);
    } catch (err) {
      notify("Failed to update profile", "error");
      handleFirestoreError(err, OperationType.UPDATE, "appUsers");
    }
  };

  const handleExportCSV = () => {
    if (leads.length === 0) return;
    const dataToExport = leads.map(l => ({
      Organisation: l.organisation || "",
      Status: l.leadStatus || "",
      ContactName: l.contactName || "",
      LeadOwner: l.leadOwner || "",
      Type: l.leadType || "",
      Email: l.email || "",
      Phone: l.phone || "",
      Location: l.location || "",
      Description: l.description || "",
      LinkedIn: l.linkedin || "",
      Website: l.website || "",
      LastContact: l.lastContactDate || "",
      PotentialDate: l.potentialProjectDate || "",
      ConfidenceScore: l.confidenceScore || "",
      StrategicAlignment: l.strategicAlignment || ""
    }));

    const csv = Papa.unparse(dataToExport);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `dcm_leads_export_${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleDownloadTemplate = () => {
    const headers = [
      "Organisation", "Activity", "Attachments", "City", "Comments", 
      "Contact Name", "Country", "Direct link", "Donation", "Email", 
      "Last Contact", "Lead Type", "Lead owner", "Lead status", 
      "Project Languages", "Project Potential date", "Referred by", 
      "Website", "What's next?"
    ];
    const csvContent = headers.join(",") + "\n";
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = "dcm_leads_template.csv";
    a.click();
    window.URL.revokeObjectURL(url);
  };

  const handleImport = async (text: string) => {
    if (!text) return;
    setLoading(true);
    try {
      Papa.parse(text, {
        header: true,
        skipEmptyLines: true,
        transformHeader: (header) => header.trim(),
        complete: async (results) => {
          const rows = results.data;
          const CHUNK_SIZE = 400;
          try {
            for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
              const chunk = rows.slice(i, i + CHUNK_SIZE);
              const batch = writeBatch(db);
              chunk.forEach((row: any) => {
                const orgName = (String(row["Organisation"] || "")).trim().substring(0, 499);
                if (!orgName) return;
                const docRef = doc(collection(db, "leads"));
                const currentStatuses = settings?.leadsConfig?.statuses || [
                  "Under assessment", "First meeting/contact", "Needs identifying", "Follow up", 
                  "Under consideration", "Approved for future project", "Strong potential", 
                  "Potential", "Not Interested", "Not suitable", "Project Completed", "Closed"
                ];
                const rawStatus = (row["Lead status"] || "").trim().toLowerCase();
                const matchedStatus = currentStatuses.find(s => s.toLowerCase() === rawStatus) || 
                                     (rawStatus.includes("closed") ? "Closed" : 
                                      rawStatus.includes("completed") ? "Project Completed" :
                                      rawStatus.includes("not interest") ? "Not Interested" :
                                      "Under assessment");

                batch.set(docRef, {
                  organisation: orgName,
                  activity: row["Activity"] || "",
                  attachments: row["Attachments"] || "",
                  city: row["City"] || "",
                  comments: row["Comments"] || "",
                  contactName: row["Contact Name"] || "",
                  country: row["Country"] || "",
                  directLink: row["Direct link"] || "",
                  donation: row["Donation"] || "",
                  email: row["Email"] || "",
                  lastContact: row["Last Contact"] || "",
                  leadType: row["Lead Type"] || "",
                  leadOwner: row["Lead owner"] || "",
                  leadStatus: matchedStatus,
                  projectLanguages: row["Project Languages"] || "",
                  projectPotentialDate: row["Project Potential date"] || "",
                  referredBy: row["Referred by"] || "",
                  website: row["Website"] || "",
                  whatsNext: row["What's next?"] || "",
                  createdAt: new Date().toISOString(),
                  updatedAt: new Date().toISOString()
                });
              });
              await batch.commit();
            }
            alert(`Imported ${rows.length} leads successfully.`);
          } catch (err) {
            console.error("Batch commit error:", err);
          } finally {
            setLoading(false);
          }
        }
      });
    } catch (err) {
      console.error("Import error:", err);
      setLoading(false);
    }
  };

  const isSystemAdmin = [
    "fredericf.fery@gmail.com", 
    "frederic@datachangemakers.org",
    "zifrench2@gmail.com"
  ].includes(auth.currentUser?.email?.toLowerCase() || "");

  return (
    <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="space-y-10">
      <div className="flex gap-4 border-b border-brand-border pb-4">
        {[
          { id: 'users', label: 'User Management', icon: Users },
          { id: 'scoring', label: 'Scoring Rules', icon: CheckCircle2 },
          { id: 'ai', label: 'AI Parameters', icon: Activity },
          { id: 'leads', label: 'Leads Config', icon: HeartHandshake },
          { id: 'permissions', label: 'Permissions', icon: ShieldCheck },
          { id: 'data', label: 'Data & Backups', icon: Download },
          { id: 'archive', label: 'Lead Archive', icon: Archive }
        ].map((t) => (
          <button
            key={t.id}
            onClick={() => setActiveAdminTab(t.id as any)}
            className={cn(
              "px-6 py-2.5 rounded-full text-[10px] font-black uppercase tracking-widest transition-all flex items-center gap-2",
              activeAdminTab === t.id ? "bg-brand-primary text-white" : "bg-white text-brand-muted border border-brand-border hover:border-brand-accent"
            )}
          >
            <t.icon className="w-3.5 h-3.5" />
            {t.label}
          </button>
        ))}
      </div>

      <AnimatePresence mode="wait">
        {activeAdminTab === "users" && (
          <motion.div key="users" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} className="space-y-8">
            <div className="bg-white border border-brand-border rounded-2xl p-8 card-shadow">
              <h4 className="text-xs font-black text-brand-primary uppercase tracking-widest mb-6">Add New User</h4>
              <form onSubmit={addUser} className="flex flex-wrap gap-4 items-end">
                <div className="flex-1 min-w-[200px] space-y-2">
                  <label className="text-[10px] font-bold text-brand-muted uppercase">First Name</label>
                  <input 
                    type="text" 
                    value={newUserFirstName}
                    onChange={(e) => setNewUserFirstName(e.target.value)}
                    placeholder="Jane"
                    className="w-full bg-brand-bg border border-brand-border rounded-xl px-4 py-2.5 text-sm outline-none focus:ring-1 focus:ring-brand-accent shadow-sm"
                  />
                </div>
                <div className="flex-1 min-w-[200px] space-y-2">
                  <label className="text-[10px] font-bold text-brand-muted uppercase">Last Name</label>
                  <input 
                    type="text" 
                    value={newUserLastName}
                    onChange={(e) => setNewUserLastName(e.target.value)}
                    placeholder="Doe"
                    className="w-full bg-brand-bg border border-brand-border rounded-xl px-4 py-2.5 text-sm outline-none focus:ring-1 focus:ring-brand-accent shadow-sm"
                  />
                </div>
                <div className="flex-1 min-w-[250px] space-y-2">
                  <label className="text-[10px] font-bold text-brand-muted uppercase">Email Address</label>
                  <input 
                    type="email" 
                    value={newUserEmail}
                    onChange={(e) => setNewUserEmail(e.target.value)}
                    placeholder="teammate@organization.org"
                    className="w-full bg-brand-bg border border-brand-border rounded-xl px-4 py-2.5 text-sm outline-none focus:ring-1 focus:ring-brand-accent shadow-sm"
                  />
                </div>
                <div className="flex-1 min-w-[200px] space-y-2">
                  <label className="text-[10px] font-bold text-brand-muted uppercase">Password (Optional)</label>
                  <input 
                    type="text" 
                    value={newUserPassword}
                    onChange={(e) => setNewUserPassword(e.target.value)}
                    placeholder="Defaults to Changemaker2026!"
                    className="w-full bg-brand-bg border border-brand-border rounded-xl px-4 py-2.5 text-sm outline-none focus:ring-1 focus:ring-brand-accent shadow-sm font-mono"
                  />
                </div>
                <div className="w-48 space-y-2">
                  <label className="text-[10px] font-bold text-brand-muted uppercase">Role</label>
                  <select 
                    value={newUserRole}
                    onChange={(e) => setNewUserRole(e.target.value)}
                    className="w-full bg-brand-bg border border-brand-border rounded-xl px-4 py-2.5 text-xs font-bold uppercase tracking-tight outline-none focus:ring-1 focus:ring-brand-accent shadow-sm appearance-none"
                  >
                    <option value="Admin">Admin</option>
                    <option value="DCM Boards">DCM Boards</option>
                    <option value="Assessor">Assessor</option>
                  </select>
                </div>
                <button 
                  type="submit"
                  disabled={sendingEmail}
                  className="px-8 py-2.5 bg-brand-accent text-white rounded-xl font-black text-[10px] uppercase tracking-widest hover:scale-105 transition-all shadow-lg shadow-brand-accent/20 disabled:opacity-50 flex items-center gap-2"
                >
                  {sendingEmail ? <Loader2 className="w-3 h-3 animate-spin"/> : null}
                  {sendingEmail ? "Inviting..." : "Confirm & Send Email"}
                </button>
              </form>
            </div>

            <AnimatePresence>
              {invitedUser && (
                <motion.div 
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: 'auto' }}
                  exit={{ opacity: 0, height: 0 }}
                  className="bg-emerald-50 border border-emerald-100 rounded-2xl p-6 overflow-hidden"
                >
                  <div className="flex items-start gap-4">
                    <div className="w-10 h-10 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-600 shrink-0">
                      <Mail className="w-5 h-5" />
                    </div>
                    <div className="flex-1 space-y-4">
                      <div>
                        <h4 className="text-sm font-black text-emerald-800 uppercase tracking-tight">User Invitation Created</h4>
                        <p className="text-[10px] font-bold text-emerald-600 uppercase tracking-widest mt-1">
                          Successfully added {invitedUser.displayName}. Now send them their invitation link.
                        </p>
                      </div>
                      
                      <div className="bg-white/50 border border-emerald-200 rounded-xl p-4 font-mono text-[10px] break-all text-emerald-700">
                        {window.location.origin}
                      </div>

                      <div className="flex gap-3">
                        <button 
                          onClick={() => {
                            const subject = encodeURIComponent("Welcome to Changemaker Systems!");
                            const body = encodeURIComponent(`Hi ${invitedUser.firstName || invitedUser.displayName},\n\nYou've been invited to join Changemaker Systems as a ${invitedUser.role}.\n\nYou can log in using your Google account here:\n${window.location.origin}\n\nBest regards,\nThe Changemaker Team`);
                            window.location.href = `mailto:${invitedUser.email}?subject=${subject}&body=${body}`;
                          }}
                          className="px-6 py-2 bg-emerald-600 text-white rounded-lg font-black text-[10px] uppercase tracking-widest hover:bg-emerald-700 transition-all flex items-center gap-2"
                        >
                          <Mail className="w-3.5 h-3.5" /> Draft Welcome Email
                        </button>
                        <button 
                          onClick={() => {
                            navigator.clipboard.writeText(window.location.origin);
                            alert("Link copied to clipboard!");
                          }}
                          className="px-6 py-2 bg-white border border-emerald-200 text-emerald-700 rounded-lg font-black text-[10px] uppercase tracking-widest hover:bg-emerald-50 transition-all flex items-center gap-2"
                        >
                          <Copy className="w-3.5 h-3.5" /> Copy App Link
                        </button>
                        <button 
                          onClick={() => setInvitedUser(null)}
                          className="px-6 py-2 text-emerald-600 font-black text-[10px] uppercase tracking-widest hover:text-emerald-800"
                        >
                          Dismiss
                        </button>
                      </div>
                    </div>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            <div className="bg-white border border-brand-border rounded-2xl overflow-hidden card-shadow">
              <div className="p-6 border-b border-brand-border bg-brand-bg/30 flex justify-between items-center">
                <h4 className="text-[10px] font-black text-brand-primary uppercase tracking-[0.2em]">Active Team Members</h4>
                <div className="relative">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-brand-muted" />
                  <input 
                    type="text" 
                    placeholder="Search users..." 
                    className="pl-9 pr-4 py-1.5 bg-white border border-brand-border rounded-lg text-[10px] font-bold outline-none focus:ring-1 focus:ring-brand-accent w-64 uppercase tracking-wider"
                    value={userSearchTerm}
                    onChange={(e) => setUserSearchTerm(e.target.value)}
                  />
                </div>
              </div>
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-brand-bg text-brand-muted font-bold text-[10px] uppercase tracking-widest border-b border-brand-border">
                    <th className="px-8 py-5">User</th>
                    <th className="px-8 py-5">Current Role</th>
                    <th className="px-8 py-5 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-brand-border">
                  {users
                    .filter(u => 
                      !userSearchTerm || 
                      (u.email || "").toLowerCase().includes(userSearchTerm.toLowerCase()) ||
                      (u.displayName || "").toLowerCase().includes(userSearchTerm.toLowerCase()) ||
                      (u.firstName || "").toLowerCase().includes(userSearchTerm.toLowerCase()) ||
                      (u.lastName || "").toLowerCase().includes(userSearchTerm.toLowerCase())
                    )
                    .filter((u, index, self) => {
                      // If this is an "Invited" user (id === email)
                      if (u.id === u.email) {
                        // Only show if there isn't an "Active" user (id !== email) with the same email
                        return !self.some(other => other.email === u.email && other.id !== other.email);
                      }
                      return true;
                    })
                    .map((u) => (
                      <tr key={u.id} className="hover:bg-brand-bg/20 transition-colors">
                      <td className="px-8 py-6">
                          <div className="flex items-center gap-3 group/user">
                            <div className="w-8 h-8 rounded-lg bg-brand-bg flex items-center justify-center text-[10px] font-black text-brand-accent">
                              {(u.firstName?.[0] || u.displayName?.[0] || "?").toUpperCase()}
                            </div>
                            <div 
                              className="cursor-pointer group-hover:bg-brand-bg/50 px-2 py-1 rounded-lg transition-all"
                              onClick={() => {
                                setEditingUser(u);
                                setIsEditingUser(true);
                              }}
                            >
                              <p className="text-sm font-black text-brand-primary group-hover:text-brand-accent transition-colors">
                                {u.firstName || u.lastName ? `${u.firstName || ""} ${u.lastName || ""}`.trim() : (u.displayName || "Unknown User")}
                              </p>
                              <p className="text-[10px] font-mono text-brand-muted">{u.email}</p>
                            </div>
                          </div>
                      </td>
                      <td className="px-8 py-6">
                  <span className={cn(
                    "px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-widest",
                    u.role?.toLowerCase() === "admin" ? "bg-purple-100 text-purple-700" :
                    u.role?.toLowerCase() === "dcm boards" ? "bg-blue-100 text-blue-700" :
                    "bg-amber-100 text-amber-700"
                  )}>
                    {u.role}
                  </span>
                  {u.id !== u.email && (
                    <span className="ml-2 px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700 text-[8px] font-bold">Active</span>
                  )}
                  {u.id === u.email && (
                    <span className="ml-2 px-2 py-0.5 rounded-full bg-blue-100 text-blue-700 text-[8px] font-bold">Invited</span>
                  )}
                      </td>
                      <td className="px-8 py-6">
                        <div className="flex justify-end gap-3 items-center">
                          {isSystemAdmin && (
                            <button 
                              onClick={() => {
                                notify(`Testing role: ${u.role}`, "success");
                                // Real impersonation logic would go here
                              }}
                              className="p-2 text-brand-muted hover:text-amber-500 hover:bg-amber-50 rounded-lg transition-all flex items-center gap-2 border border-transparent hover:border-amber-200"
                              title="Review user access"
                            >
                              <UserCircle className="w-3.5 h-3.5" />
                              <span className="text-[10px] font-black uppercase tracking-widest pr-1">Impersonate</span>
                            </button>
                          )}
                          {u.id === u.email ? (
                            <div className="flex items-center gap-1.5 bg-amber-50/80 text-amber-800 border border-amber-200/50 rounded-lg px-2.5 py-1 text-[9px] font-bold uppercase tracking-widest" title="This pre-authorized user has not signed up yet. Editing their password in User Editor resets their initial registration password.">
                              <span className="w-1.5 h-1.5 rounded-full bg-amber-500 animate-pulse shrink-0"></span> Invited (Awaiting Sign-in)
                            </div>
                          ) : (
                            <button 
                              onClick={async () => {
                                try {
                                  await sendPasswordResetEmail(auth, u.email);
                                  notify(`Direct password reset link successfully sent to ${u.email}!`, "success");
                                } catch (err: any) {
                                  notify(err.message || "Failed to send reset link.", "error");
                                }
                              }}
                              className="p-2 text-brand-muted hover:text-brand-accent hover:bg-brand-bg rounded-lg transition-all flex items-center gap-2 border border-transparent hover:border-brand-accent/20"
                              title="Send official password reset link email to this registered user"
                            >
                              <Lock className="w-3.5 h-3.5" />
                              <span className="text-[10px] font-black uppercase tracking-widest pr-1">Reset PW</span>
                            </button>
                          )}
                          <button 
                            onClick={() => {
                              setEditingUser(u);
                              setIsEditingUser(true);
                            }}
                            className="p-2 text-brand-muted hover:text-brand-accent hover:bg-brand-bg rounded-lg transition-all flex items-center gap-2 border border-transparent hover:border-brand-accent/20"
                            title="Edit User Details"
                          >
                            <Edit3 className="w-3.5 h-3.5" />
                            <span className="text-[10px] font-black uppercase tracking-widest pr-1">Edit</span>
                          </button>
                          <select 
                            value={u.role}
                            onChange={(e) => updateUserRole(u.id, e.target.value)}
                            className="bg-brand-bg border border-brand-border rounded-lg px-3 py-1.5 text-[10px] font-bold uppercase tracking-tight outline-none focus:ring-1 focus:ring-brand-accent appearance-none cursor-pointer"
                          >
                            <option value="Admin">Admin</option>
                            <option value="DCM Boards">DCM Boards</option>
                            <option value="Assessor">Assessor</option>
                          </select>
                          <button 
                            onClick={() => deleteUser(u.id)}
                            className={cn(
                              "p-1.5 rounded-md transition-all flex items-center justify-center",
                              confirmDeleteId === u.id 
                                ? "bg-red-500 text-white" 
                                : "text-brand-muted hover:text-red-500 hover:bg-red-50"
                            )}
                            title={confirmDeleteId === u.id ? "Click again to confirm" : "Remove User"}
                          >
                            {confirmDeleteId === u.id ? (
                              <div className="flex items-center gap-1">
                                <Trash2 className="w-4 h-4" />
                                <span className="text-[8px] font-black uppercase">Confirm?</span>
                              </div>
                            ) : (
                              <Trash2 className="w-4 h-4" />
                            )}
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {isEditingUser && editingUser && (
              <div className="fixed inset-0 bg-brand-primary/60 backdrop-blur-sm z-50 flex items-center justify-center p-6">
                <motion.div 
                  initial={{ opacity: 0, scale: 0.95 }} 
                  animate={{ opacity: 1, scale: 1 }} 
                  className="bg-white rounded-3xl w-full max-w-lg shadow-2xl overflow-hidden border border-brand-border"
                >
                  <div className="p-8 border-b border-brand-border bg-brand-bg flex justify-between items-center">
                    <div>
                      <h3 className="text-xl font-black text-brand-primary uppercase tracking-tight">Edit User</h3>
                      <p className="text-[10px] font-bold text-brand-muted uppercase tracking-widest mt-1">Modify account details and permissions</p>
                    </div>
                    <button onClick={() => setIsEditingUser(false)} className="p-2 hover:bg-brand-border rounded-full transition-colors text-brand-muted">
                      <X className="w-6 h-6" />
                    </button>
                  </div>
                  <form onSubmit={handleUpdateUser} className="p-8 space-y-6">
                    <div className="grid grid-cols-2 gap-4">
                      <div className="space-y-2">
                        <label className="text-[10px] font-bold text-brand-muted uppercase">First Name</label>
                        <input 
                          className="w-full bg-brand-bg border border-brand-border rounded-xl px-4 py-2.5 text-sm outline-none focus:ring-1 focus:ring-brand-accent shadow-sm" 
                          value={editingUser.firstName || ''} 
                          onChange={e => setEditingUser({...editingUser, firstName: e.target.value})} 
                        />
                      </div>
                      <div className="space-y-2">
                        <label className="text-[10px] font-bold text-brand-muted uppercase">Last Name</label>
                        <input 
                          className="w-full bg-brand-bg border border-brand-border rounded-xl px-4 py-2.5 text-sm outline-none focus:ring-1 focus:ring-brand-accent shadow-sm" 
                          value={editingUser.lastName || ''} 
                          onChange={e => setEditingUser({...editingUser, lastName: e.target.value})} 
                        />
                      </div>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-bold text-brand-muted uppercase">Email Address</label>
                      <input 
                        type="email" 
                        className="w-full bg-brand-bg border border-brand-border rounded-xl px-4 py-2.5 text-sm outline-none focus:ring-1 focus:ring-brand-accent shadow-sm" 
                        value={editingUser.email || ''} 
                        onChange={e => setEditingUser({...editingUser, email: e.target.value})} 
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-bold text-brand-muted uppercase">User Role</label>
                      <select 
                        className="w-full bg-brand-bg border border-brand-border rounded-xl px-4 py-2.5 text-sm outline-none focus:ring-1 focus:ring-brand-accent shadow-sm cursor-pointer" 
                        value={editingUser.role} 
                        onChange={e => setEditingUser({...editingUser, role: e.target.value})}
                      >
                        <option value="Admin">Admin</option>
                        <option value="DCM Boards">DCM Boards</option>
                        <option value="Assessor">Assessor</option>
                      </select>
                    </div>
                    <div className="space-y-4">
                      <div className="space-y-2">
                        <label className="text-[10px] font-bold text-brand-muted uppercase">User Password</label>
                        <input 
                          type="text" 
                          placeholder="e.g. Changemaker2026! or a custom password"
                          className="w-full bg-brand-bg border border-brand-border rounded-xl px-4 py-2.5 text-sm outline-none focus:ring-1 focus:ring-brand-accent shadow-sm font-mono" 
                          value={editingUser.password || ''} 
                          onChange={e => setEditingUser({...editingUser, password: e.target.value})} 
                        />
                        <p className="text-[10px] text-brand-muted pl-1 leading-relaxed">
                          {editingUser.id === editingUser.email 
                            ? "This custom initial password is used for their first-time login authorization." 
                            : "This local field stores their default password. Since this user is already active, updating this text alone will not update their Firebase Authenticated password."}
                        </p>
                      </div>

                      {editingUser.id !== editingUser.email && (
                        <div className="bg-amber-50 border border-amber-200/50 rounded-xl p-4 space-y-3">
                          <div className="flex gap-2.5 items-start text-amber-800 text-xs font-semibold">
                            <AlertCircle className="w-4 h-4 shrink-0 text-amber-600 mt-0.5" />
                            <div>
                              <p className="font-bold">Active User Reset Guideline</p>
                              <p className="text-[10px] text-amber-700/80 mt-1 font-normal leading-relaxed">
                                Because of security regulations, to configure a new password for active teammates, please use the secure button below. It will send an official secure Firebase reset email to this user.
                              </p>
                            </div>
                          </div>
                          <button
                            type="button"
                            onClick={async () => {
                              try {
                                await sendPasswordResetEmail(auth, editingUser.email);
                                notify(`Direct password reset link successfully sent to ${editingUser.email}!`, "success");
                              } catch (err: any) {
                                notify(err.message || "Failed to send reset link.", "error");
                              }
                            }}
                            className="w-full flex items-center justify-center gap-2 bg-white hover:bg-amber-100/50 border border-amber-200 text-amber-800 font-bold uppercase tracking-widest text-[9px] py-2.5 px-4 rounded-xl transition-all shadow-sm"
                          >
                            <Mail className="w-3.5 h-3.5" /> Send Password Reset Email
                          </button>
                        </div>
                      )}
                    </div>
                    <div className="flex justify-end gap-3 pt-6 border-t border-brand-border">
                      <button type="button" onClick={() => setIsEditingUser(false)} className="px-6 py-2.5 font-bold text-brand-muted uppercase tracking-widest text-[10px] hover:text-brand-primary">Cancel</button>
                      <button type="submit" className="px-10 py-3 bg-brand-primary text-white rounded-xl font-black text-[10px] uppercase tracking-widest hover:scale-[1.02] active:scale-[0.98] transition-all shadow-lg shadow-brand-primary/20">Save Changes</button>
                    </div>
                  </form>
                </motion.div>
              </div>
            )}
          </motion.div>
        )}

        {activeAdminTab === "scoring" && (
          <motion.div key="scoring" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} className="space-y-12">
            {['verification', 'validation', 'validationChecks'].map((sectionKey) => (
              <div key={sectionKey} className="space-y-6">
                <div className="flex justify-between items-end">
                  <div>
                    <h4 className="text-lg font-black text-brand-primary uppercase tracking-tight">{sectionKey.replace(/([A-Z])/g, ' $1')} Questions</h4>
                    <p className="text-xs text-brand-muted font-medium">Manage criteria for this evaluation pillar</p>
                  </div>
                  <button 
                    onClick={() => {
                      const newId = `q_${Date.now()}`;
                      const newList = [...(settings?.scoring?.[sectionKey] || []), { id: newId, label: "New Question", type: sectionKey === 'validation' ? 'score' : 'boolean' }];
                      updateScoringConfig(sectionKey, newList);
                    }}
                    className="flex items-center gap-2 bg-brand-accent text-white px-4 py-2 rounded-lg text-[10px] font-black uppercase tracking-widest shadow-sm hover:scale-105 transition-all"
                  >
                    <Plus className="w-3.5 h-3.5" /> ADD QUESTION
                  </button>
                </div>
                
                <div className="grid gap-4">
                  {settings?.scoring?.[sectionKey]?.map((q: any, idx: number) => (
                    <div key={q.id} className="bg-white border border-brand-border p-6 rounded-2xl card-shadow flex gap-6 items-start">
                      <div className="w-8 h-8 rounded-full bg-brand-bg flex items-center justify-center text-[10px] font-black text-brand-muted shrink-0">
                        {idx + 1}
                      </div>
                      <div className="flex-1 space-y-4">
                        <textarea 
                          value={q.label}
                          onChange={(e) => {
                            const newList = settings.scoring[sectionKey].map((item: any) => item.id === q.id ? { ...item, label: e.target.value } : item);
                            updateScoringConfig(sectionKey, newList);
                          }}
                          className="w-full bg-brand-bg/50 border border-brand-border rounded-xl p-3 text-xs font-bold outline-none focus:ring-1 focus:ring-brand-accent resize-none min-h-[60px]"
                        />
                        <div className="flex items-center gap-6">
                          <div className="flex items-center gap-2">
                             <span className="text-[9px] font-black text-brand-muted uppercase">Type:</span>
                             <select 
                               value={q.type}
                               onChange={(e) => {
                                 const newList = settings.scoring[sectionKey].map((item: any) => item.id === q.id ? { ...item, type: e.target.value } : item);
                                 updateScoringConfig(sectionKey, newList);
                               }}
                               className="bg-transparent text-[10px] font-bold text-brand-accent uppercase outline-none"
                             >
                               <option value="boolean">Boolean (Yes/No)</option>
                               <option value="score">Numerical (0-4)</option>
                             </select>
                          </div>
                          <div className="flex-1 flex items-center gap-2">
                             <span className="text-[9px] font-black text-brand-muted uppercase whitespace-nowrap">Hint:</span>
                             <input 
                               type="text"
                               placeholder="Optional hint..."
                               value={q.hint || ""}
                               onChange={(e) => {
                                 const newList = settings.scoring[sectionKey].map((item: any) => item.id === q.id ? { ...item, hint: e.target.value } : item);
                                 updateScoringConfig(sectionKey, newList);
                               }}
                               className="flex-1 bg-transparent border-b border-brand-border py-1 text-[10px] font-medium outline-none focus:border-brand-accent transition-colors"
                             />
                          </div>
                        </div>
                      </div>
                      <button 
                        onClick={() => {
                          const newList = settings.scoring[sectionKey].filter((item: any) => item.id !== q.id);
                          updateScoringConfig(sectionKey, newList);
                        }}
                        className="p-2 text-brand-muted hover:text-red-500 hover:bg-red-50 rounded-lg transition-all"
                      >
                        <Trash2 className="w-5 h-5" />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </motion.div>
        )}

        {activeAdminTab === "ai" && (
          <motion.div key="ai" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} className="max-w-2xl space-y-10">
            {/* Connection Status Section */}
            <div className="bg-white border border-brand-border rounded-2xl p-8 card-shadow">
              <div className="flex items-center gap-3 mb-6">
                <div className="w-10 h-10 rounded-xl bg-brand-bg flex items-center justify-center text-brand-primary">
                  <ShieldCheck className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-sm font-black text-brand-primary uppercase tracking-wider">System Integration Status</h3>
                  <p className="text-xs text-brand-muted">Verify connectivity to AI and external data services</p>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="bg-brand-bg/50 border border-brand-border rounded-xl p-4 flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    <div className="w-8 h-8 rounded-lg bg-blue-500/10 flex items-center justify-center text-blue-600">
                      <Cpu className="w-4 h-4" />
                    </div>
                    <div>
                      <p className="text-[10px] font-black uppercase tracking-widest text-brand-muted">Gemini AI Service</p>
                      <p className="text-[11px] font-bold text-brand-primary mt-0.5">@google/genai (Flash 1.5)</p>
                    </div>
                  </div>
                  <button 
                    onClick={async () => {
                      try {
                        const customKey = sessionApiKey || settings?.aiResearch?.apiKey || "";
                        const testKey = async (key: string) => {
                          try {
                            const ai = new GoogleGenAI({ apiKey: key });
                            const result = await ai.models.generateContent({
                              model: "gemini-3-flash-preview",
                              contents: "ping",
                            });
                            return !!result.text;
                          } catch (err) {
                            return false;
                          }
                        };
                        const isValid = await testKey(customKey);
                        if (isValid) {
                          notify(`AI Online (${customKey ? customKey.substring(0, 8) + '...' : 'Env'})`, "success");
                        } else if (customKey) {
                          notify("AI Key Invalid or Quota reached", "error");
                        } else {
                          notify("AI Key Missing (Env/Settings)", "error");
                        }
                      } catch (e) {
                        notify("AI Status Unknown", "error");
                      }
                    }}
                    className="px-3 py-1.5 bg-white border border-brand-border rounded-lg text-[10px] font-black uppercase tracking-widest hover:bg-brand-bg transition-all"
                  >
                    Verify Link
                  </button>
                </div>

                <div className="bg-brand-bg/50 border border-brand-border rounded-xl p-4 flex items-center justify-between opacity-60">
                  <div className="flex items-center gap-3">
                    <div className="w-8 h-8 rounded-lg bg-orange-500/10 flex items-center justify-center text-orange-600">
                      <ExternalLink className="w-4 h-4" />
                    </div>
                    <div>
                      <p className="text-[10px] font-black uppercase tracking-widest text-brand-muted">ProPublica API</p>
                      <p className="text-[11px] font-bold text-brand-primary mt-0.5">Non-profit Financial Data</p>
                    </div>
                  </div>
                  <span className="text-[9px] font-black text-emerald-500 uppercase tracking-widest">Connected</span>
                </div>
              </div>
            </div>

            <div className="bg-white border border-brand-border rounded-2xl p-8 card-shadow space-y-8">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="text-lg font-black text-brand-primary uppercase tracking-tight mb-2">Intelligence Control Hub</h4>
                  <p className="text-xs text-brand-muted font-medium">Fine-tune AI agent behavior and target intelligence vectors</p>
                </div>
                <div className="flex flex-col items-end">
                   <span className="text-[10px] font-black uppercase text-brand-muted">API Health</span>
                   {apiStatus?.isValid ? (
                     <span className="text-[10px] font-black uppercase text-emerald-600 flex items-center gap-1">
                        <Check className="w-3 h-3" /> Online
                     </span>
                   ) : apiStatus?.error ? (
                     <span className="text-[10px] font-black uppercase text-rose-600 flex items-center gap-1">
                        <X className="w-3 h-3" /> Error
                     </span>
                   ) : (
                     <span className="text-[10px] font-black uppercase text-brand-muted">Not Tested</span>
                   )}
                </div>
              </div>

              <div className="space-y-6">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-3">
                    <label className="text-[10px] font-black text-brand-primary uppercase tracking-widest">Gemini Project API Key</label>
                    <div className="relative">
                      <input 
                        type="password"
                        placeholder="Paste your Gemini API key here..."
                        value={settings?.aiResearch?.apiKey || ""}
                        onChange={(e) => updateAIConfig({ ...settings.aiResearch, apiKey: e.target.value })}
                        className="w-full bg-brand-bg border border-brand-border rounded-xl p-4 text-xs font-bold outline-none focus:ring-1 focus:ring-brand-accent shadow-sm"
                      />
                      <Database className="absolute right-4 top-1/2 -translate-y-1/2 w-4 h-4 text-brand-muted" />
                    </div>
                    <div className="flex justify-between items-center">
                      <p className="text-[9px] font-bold text-brand-muted uppercase">Overrides environment config.</p>
                      <button 
                        onClick={async () => {
                          if (isCheckInProgress) return;
                          setIsCheckInProgress(true);
                          try {
                            const customKey = settings?.aiResearch?.apiKey || "";
                            const ai = new GoogleGenAI({ apiKey: customKey });
                            const result = await ai.models.generateContent({
                              model: "gemini-3-flash-preview",
                              contents: "ping",
                            });
                            const isValid = !!result.text;
                            setApiStatus({ 
                              isValid, 
                              hasKey: true, 
                              keyPrefix: customKey.substring(0, 8) + "..." 
                            });
                            if (isValid) notify("Key verified successfully!", "success");
                            else notify("Verification failed: No text response", "error");
                          } catch (err: any) {
                            setApiStatus({ isValid: false, error: err.message || "Invalid Key" });
                            notify(`Verification failed: ${err.message}`, "error");
                          } finally {
                            setIsCheckInProgress(false);
                          }
                        }}
                        disabled={isCheckInProgress}
                        className="text-[9px] font-black uppercase text-brand-accent hover:underline disabled:opacity-50"
                      >
                        {isCheckInProgress ? "Verifying..." : "Test Keys Now"}
                      </button>
                    </div>
                  </div>

                  <div className="space-y-3">
                    <label className="text-[10px] font-black text-brand-primary uppercase tracking-widest">Active Intelligence Engine</label>
                    <select 
                      value={settings?.aiResearch?.model || "gemini-1.5-flash"}
                      onChange={(e) => updateAIConfig({ ...settings.aiResearch, model: e.target.value })}
                      className="w-full bg-brand-bg border border-brand-border rounded-xl p-4 text-xs font-bold outline-none focus:ring-1 focus:ring-brand-accent shadow-sm appearance-none cursor-pointer"
                    >
                      <optgroup label="Stable Models (Recommended)">
                        <option value="gemini-1.5-flash">Gemini 1.5 Flash (Fast)</option>
                        <option value="gemini-1.5-pro">Gemini 1.5 Pro (Powerful)</option>
                      </optgroup>
                      <optgroup label="Advanced (Intelligence Models)">
                        <option value="gemini-3-flash-preview">Gemini 3 Flash (Latest)</option>
                        <option value="gemini-3.1-pro-preview">Gemini 3.1 Pro (Heavy Tasks)</option>
                      </optgroup>
                    </select>
                    <p className="text-[9px] font-bold text-brand-muted uppercase">Flash models recommended for volume research.</p>
                  </div>
                </div>

                {apiStatus?.error && (
                  <div className="p-4 bg-rose-50 border border-rose-100 rounded-xl">
                    <div className="flex items-center gap-2 mb-1">
                      <AlertCircle className="w-3.5 h-3.5 text-rose-600" />
                      <span className="text-[10px] font-black uppercase text-rose-800">Connection Details</span>
                    </div>
                    <p className="text-[10px] text-rose-600 font-mono break-all leading-tight">
                      {apiStatus.error}
                    </p>
                    <p className="text-[9px] text-rose-500 mt-2 italic">
                      Verify your RPM (Requests Per Minute) and RPD (Requests Per Day) limits in Google AI Studio.
                    </p>
                  </div>
                )}

                <div className="space-y-3">
                  <div className="flex justify-between items-center">
                    <label className="text-[10px] font-black text-brand-primary uppercase tracking-widest">Deep Research Sources & Target URLs</label>
                  </div>
                  <p className="text-[9px] font-medium text-brand-muted mt-1 uppercase leading-tight">
                    Adding URLs here forces the AI agent to prioritize scraping and auditing those specific digital assets.
                  </p>
                  
                  <div className="flex gap-2">
                    <input 
                      type="text" 
                      placeholder="Enter source name or URL..."
                      className="flex-1 bg-brand-bg border border-brand-border rounded-xl px-4 py-2 text-[10px] font-bold outline-none focus:ring-1 focus:ring-brand-accent shadow-sm"
                      value={newSource}
                      onChange={(e) => setNewSource(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' && newSource.trim()) {
                          updateAIConfig({ ...settings.aiResearch, sources: [...(settings.aiResearch.sources || []), newSource.trim()] });
                          setNewSource("");
                        }
                      }}
                    />
                    <button 
                      onClick={() => {
                        if (newSource.trim()) {
                          updateAIConfig({ ...settings.aiResearch, sources: [...(settings.aiResearch.sources || []), newSource.trim()] });
                          setNewSource("");
                        }
                      }}
                      className="px-4 py-2 bg-brand-accent text-white text-[10px] font-black uppercase rounded-xl hover:bg-brand-accent/90 transition-all shadow-sm"
                    >
                      Add
                    </button>
                  </div>

                  <div className="flex flex-wrap gap-2 pt-2">
                    {settings?.aiResearch?.sources?.map((s: string) => (
                      <span key={s} className="bg-brand-bg border border-brand-border px-3 py-1.5 rounded-full text-[10px] font-bold text-brand-primary flex items-center gap-2 group shadow-sm">
                        {s.startsWith('http') ? <ExternalLink className="w-2.5 h-2.5 text-brand-accent/70" /> : <SearchIcon className="w-2.5 h-2.5 text-brand-muted" />}
                        <span className="max-w-[150px] truncate">{s}</span>
                        <button 
                          onClick={() => updateAIConfig({ ...settings.aiResearch, sources: settings.aiResearch.sources.filter((item: string) => item !== s) })}
                          className="text-brand-muted hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity"
                        >
                          <X className="w-3 h-3" />
                        </button>
                      </span>
                    ))}
                  </div>
                </div>
              </div>
            </div>

            <div className="bg-white border border-brand-border rounded-2xl p-8 card-shadow space-y-8">
              <div>
                <h4 className="text-lg font-black text-brand-primary uppercase tracking-tight mb-2">Research Hygiene & Purge Tools</h4>
                <p className="text-xs text-brand-muted font-medium">Reset AI research records and clear hallucinations</p>
              </div>

              <div className="space-y-6">
                <div className="p-6 bg-brand-bg border border-brand-border rounded-2xl space-y-4">
                  <label className="text-[10px] font-black text-brand-primary uppercase tracking-widest">Wipe Single Lead Research</label>
                  <div className="flex gap-3">
                    <select 
                      value={selectedPurgeLeadId}
                      onChange={(e) => setSelectedPurgeLeadId(e.target.value)}
                      className="flex-1 bg-white border border-brand-border rounded-xl px-4 py-2 text-xs font-bold outline-none focus:ring-1 focus:ring-brand-accent appearance-none cursor-pointer"
                    >
                      <option value="">Select a lead to purge...</option>
                      {leads.map(l => (
                        <option key={l.id} value={l.id}>{l.organisation}</option>
                      ))}
                    </select>
                    <button 
                      onClick={() => purgeLeadResearch(selectedPurgeLeadId)}
                      disabled={!selectedPurgeLeadId || loading}
                      className={cn(
                        "px-6 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all whitespace-nowrap",
                        confirmPurgeId === selectedPurgeLeadId && selectedPurgeLeadId ? "bg-rose-600 text-white" : "bg-rose-50 text-rose-600 border border-rose-200 hover:bg-rose-100",
                        (!selectedPurgeLeadId || loading) && "opacity-50"
                      )}
                    >
                      {confirmPurgeId === selectedPurgeLeadId && selectedPurgeLeadId ? "Confirm Wash?" : "Purge Lead"}
                    </button>
                  </div>
                </div>

                <div className="p-6 bg-rose-50 border border-rose-200 rounded-2xl flex items-center justify-between gap-6">
                  <div className="space-y-1">
                    <h5 className="font-black text-xs uppercase tracking-widest text-rose-800">Global AI Reset</h5>
                    <p className="text-[10px] text-rose-600 font-medium max-w-xs">
                      Permanent WIPE of all AI-generated fields for all {leads.length} leads.
                    </p>
                  </div>
                  <button 
                    onClick={purgeAllAIResearch}
                    disabled={loading}
                    className={cn(
                      "px-6 py-3 rounded-xl font-black text-[10px] uppercase tracking-widest transition-all shadow-lg",
                      confirmGlobalPurge ? "bg-rose-900 text-white animate-pulse" : "bg-rose-600 text-white hover:bg-rose-700 shadow-rose-600/20",
                      loading && "opacity-50"
                    )}
                  >
                    {confirmGlobalPurge ? "CONFIRM GLOBAL WIPE" : "Purge ALL Research"}
                  </button>
                </div>
              </div>
            </div>
          </motion.div>
        )}
        {activeAdminTab === "leads" && (
          <motion.div key="leads" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} className="space-y-10">
            <div className="max-w-4xl grid grid-cols-1 md:grid-cols-2 gap-8">
              <div className="bg-white border border-brand-border rounded-2xl p-8 card-shadow space-y-6">
                <div className="flex justify-between items-center">
                  <div>
                    <h4 className="text-sm font-black text-brand-primary uppercase tracking-tight">Status Pipeline</h4>
                    <p className="text-[10px] text-brand-muted font-medium">Reorder, rename, or resize columns</p>
                  </div>
                  <div className="flex gap-2">
                    <button 
                      onClick={() => {
                        if (confirm("Restore standard Status Pipeline?")) {
                          updateLeadsConfig("statuses", [
                            "Under assessment",
                            "First Meeting",
                            "Needs Identifying",
                            "Follow Up",
                            "Under Consideration",
                            "Approved",
                            "Strong Potential",
                            "Potential",
                            "Not Interested",
                            "Not Suitable",
                            "Project Completed",
                            "Closed"
                          ]);
                        }
                      }}
                      className="px-2 py-1 bg-brand-bg border border-brand-border text-[8px] font-black uppercase text-brand-muted rounded hover:bg-brand-border transition-all"
                    >
                      Reset
                    </button>
                    <button 
                      onClick={() => {
                        if (newStatus.trim()) {
                          const current = settings?.leadsConfig?.statuses || [];
                          updateLeadsConfig("statuses", [...current, newStatus.trim()]);
                          setNewStatus("");
                        }
                      }}
                      className="p-1.5 bg-brand-accent/10 text-brand-accent rounded hover:bg-brand-accent/20 transition-all"
                    >
                      <Plus className="w-4 h-4" />
                    </button>
                  </div>
                </div>
                
                <div className="px-1">
                  <div className="flex gap-2 mb-4">
                    <input 
                      type="text" 
                      placeholder="Add status..."
                      className="flex-1 bg-brand-bg border border-brand-border rounded-lg px-3 py-1.5 text-[10px] font-bold outline-none focus:ring-1 focus:ring-brand-accent"
                      value={newStatus}
                      onChange={(e) => setNewStatus(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' && newStatus.trim()) {
                          const current = settings?.leadsConfig?.statuses || [];
                          updateLeadsConfig("statuses", [...current, newStatus.trim()]);
                          setNewStatus("");
                        }
                      }}
                    />
                  </div>
                </div>

                <DndContext 
                  sensors={sensors}
                  collisionDetection={closestCenter}
                  onDragEnd={(event: DragEndEvent) => {
                    const { active, over } = event;
                    if (over && active.id !== over.id) {
                      const current = settings?.leadsConfig?.statuses || [];
                      const oldIndex = current.indexOf(active.id as string);
                      const newIndex = current.indexOf(over.id as string);
                      const newList = arrayMove(current, oldIndex, newIndex);
                      updateLeadsConfig("statuses", newList);
                    }
                  }}
                >
                  <SortableContext 
                    items={settings?.leadsConfig?.statuses || []}
                    strategy={verticalListSortingStrategy}
                  >
                    <div className="space-y-2 max-h-[450px] overflow-y-auto overflow-x-hidden px-1 custom-scrollbar">
                      {(settings?.leadsConfig?.statuses || []).map((s: string) => (
                        <SortableStatusItem 
                          key={s} 
                          id={s} 
                          readOnly={false}
                          boardStatuses={settings?.leadsConfig?.statuses || []}
                          setBoardStatuses={(val: any) => {
                            if (typeof val === 'function') {
                              updateLeadsConfig("statuses", val(settings?.leadsConfig?.statuses || []));
                            } else {
                              updateLeadsConfig("statuses", val);
                            }
                          }}
                          columnWidths={settings?.leadsConfig?.columnWidths || {}}
                          setColumnWidths={(val: any) => {
                            if (typeof val === 'function') {
                              updateLeadsConfig("columnWidths", val(settings?.leadsConfig?.columnWidths || {}));
                            } else {
                              updateLeadsConfig("columnWidths", val);
                            }
                          }}
                          onUpdateGlobal={(newList) => updateLeadsConfig("statuses", newList)}
                        />
                      ))}
                    </div>
                  </SortableContext>
                </DndContext>
            </div>

            <div className="bg-white border border-brand-border rounded-2xl p-8 card-shadow space-y-6">
                <div className="flex justify-between items-center">
                  <div>
                    <h4 className="text-sm font-black text-brand-primary uppercase tracking-tight">Lead Types</h4>
                    <p className="text-[10px] text-brand-muted font-medium">Manage categories for your leads</p>
                  </div>
                  <div className="flex gap-2">
                    <button 
                      onClick={() => {
                        if (confirm("Restore standard Lead Types (Education, Partnership, Project)?")) {
                          updateLeadsConfig("types", ["Education", "Partnership", "Project"]);
                        }
                      }}
                      className="px-2 py-1 bg-brand-bg border border-brand-border text-[8px] font-black uppercase text-brand-muted rounded hover:bg-brand-border transition-all"
                    >
                      Reset
                    </button>
                    <button 
                      onClick={() => {
                        if (newLeadType.trim()) {
                          updateLeadsConfig("types", [...(settings?.leadsConfig?.types || []), newLeadType.trim()]);
                          setNewLeadType("");
                        }
                      }}
                      className="p-1.5 bg-brand-accent/10 text-brand-accent rounded hover:bg-brand-accent/20 transition-all"
                    >
                      <Plus className="w-4 h-4" />
                    </button>
                  </div>
                </div>

                <div className="px-1">
                  <div className="flex gap-2 mb-4">
                    <input 
                      type="text" 
                      placeholder="Add type..."
                      className="flex-1 bg-brand-bg border border-brand-border rounded-lg px-3 py-1.5 text-[10px] font-bold outline-none focus:ring-1 focus:ring-brand-accent"
                      value={newLeadType}
                      onChange={(e) => setNewLeadType(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' && newLeadType.trim()) {
                          updateLeadsConfig("types", [...(settings?.leadsConfig?.types || []), newLeadType.trim()]);
                          setNewLeadType("");
                        }
                      }}
                    />
                  </div>
                </div>

              <div className="space-y-2 max-h-[400px] overflow-y-auto px-1">
                {(settings?.leadsConfig?.types || []).map((t: string, idx: number) => (
                  <div key={t} className="flex items-center gap-2 p-3 bg-brand-bg border border-brand-border rounded-xl group">
                    <span className="text-[10px] font-black text-brand-muted tracking-tighter w-4">{idx + 1}</span>
                    <input 
                      className="flex-1 bg-transparent border-none text-[10px] font-bold outline-none text-brand-primary"
                      value={t}
                      onChange={(e) => {
                        const next = [...(settings?.leadsConfig?.types || [])];
                        next[idx] = e.target.value;
                        updateLeadsConfig("types", next);
                      }}
                    />
                    <button 
                      onClick={() => updateLeadsConfig("types", settings.leadsConfig.types.filter((_: any, i: number) => i !== idx))}
                      className="text-brand-muted hover:text-red-500 opacity-0 group-hover:opacity-100 transition-all p-1"
                    >
                      <Trash2 className="w-3 h-3" />
                    </button>
                  </div>
                ))}
                {(!settings?.leadsConfig?.types || settings.leadsConfig.types.length === 0) && (
                  <p className="text-[10px] text-brand-muted italic py-4 text-center">No types configured yet.</p>
                )}
              </div>
            </div>

            <div className="bg-white border border-brand-border rounded-2xl p-8 card-shadow space-y-6">
              <div className="flex justify-between items-center">
                <div>
                  <h4 className="text-sm font-black text-brand-primary uppercase tracking-tight">Display Settings</h4>
                  <p className="text-[10px] text-brand-muted font-medium">Configure date formatting preferences</p>
                </div>
              </div>
              <div className="space-y-4">
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-brand-primary uppercase tracking-widest">Date Format</label>
                  <select 
                    value={settings?.displayConfig?.dateFormat || "dd/mm/yyyy"}
                    onChange={(e) => updateDisplayConfig({ ...settings.displayConfig, dateFormat: e.target.value })}
                    className="w-full bg-brand-bg border border-brand-border rounded-xl px-4 py-2 text-xs font-bold outline-none focus:ring-1 focus:ring-brand-accent shadow-sm appearance-none cursor-pointer"
                  >
                    <option value="dd/mm/yyyy">DD/MM/YYYY (Day/Month/Year)</option>
                    <option value="mm/dd/yyyy">MM/DD/YYYY (Month/Day/Year)</option>
                    <option value="MMMM D, YYYY">Long Format (e.g. April 27, 2026)</option>
                  </select>
                </div>
              </div>
            </div>
          </div>
        </motion.div>
        )}

        {activeAdminTab === "data" && (
          <motion.div key="data" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} className="space-y-10">
            <div className="max-w-4xl bg-white border border-brand-border rounded-2xl p-10 card-shadow space-y-8 text-brand-primary">
              <div>
                <h4 className="text-xl font-black uppercase tracking-tight mb-2">System Data Management</h4>
                <p className="text-sm text-brand-muted font-medium">Export, import, and backup your DCM Lead intelligence records.</p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="bg-brand-bg border border-brand-border rounded-2xl p-6 space-y-4 hover:border-brand-accent transition-all group">
                  <div className="w-12 h-12 rounded-xl bg-brand-accent text-white flex items-center justify-center shadow-lg shadow-brand-accent/20">
                    <Download className="w-6 h-6" />
                  </div>
                  <div>
                    <h5 className="font-black text-xs uppercase tracking-widest mb-1">Export Data</h5>
                    <p className="text-[10px] text-brand-muted font-medium mb-4">Download all current leads as a CSV file for archival or external analysis.</p>
                    <button 
                      onClick={handleExportCSV}
                      className="w-full py-2.5 bg-brand-primary text-white rounded-lg text-[10px] font-black uppercase tracking-widest hover:opacity-90 transition-all"
                    >
                      Export CSV
                    </button>
                  </div>
                </div>

                <div className="bg-brand-bg border border-brand-border rounded-2xl p-6 space-y-4 hover:border-brand-accent transition-all group">
                  <div className="w-12 h-12 rounded-xl bg-blue-600 text-white flex items-center justify-center shadow-lg shadow-blue-600/20">
                    <FileText className="w-6 h-6" />
                  </div>
                  <div>
                    <h5 className="font-black text-xs uppercase tracking-widest mb-1">CSV Template</h5>
                    <p className="text-[10px] text-brand-muted font-medium mb-4">Download a clean template with the correct headers for successful importing.</p>
                    <button 
                      onClick={handleDownloadTemplate}
                      className="w-full py-2.5 bg-blue-600 text-white rounded-lg text-[10px] font-black uppercase tracking-widest hover:opacity-90 transition-all"
                    >
                      Download Template
                    </button>
                  </div>
                </div>

                <div className="bg-brand-bg border border-brand-border rounded-2xl p-6 space-y-4 hover:border-brand-accent transition-all group border-dashed relative">
                  <input
                    type="file"
                    accept=".csv"
                    className="absolute inset-0 opacity-0 cursor-pointer z-10"
                    onChange={async (e) => {
                      const file = e.target.files?.[0];
                      if (file) {
                        const text = await file.text();
                        handleImport(text);
                      }
                    }}
                  />
                  <div className="w-12 h-12 rounded-xl bg-emerald-600 text-white flex items-center justify-center shadow-lg shadow-emerald-600/20">
                    <UploadCloud className="w-6 h-6" />
                  </div>
                  <div>
                    <h5 className="font-black text-xs uppercase tracking-widest mb-1">Import Data</h5>
                    <p className="text-[10px] text-brand-muted font-medium mb-4">Upload your completed CSV file to bulk-import new leads into the system.</p>
                    <div className="w-full py-2.5 bg-emerald-600 text-white rounded-lg text-[10px] font-black uppercase tracking-widest hover:opacity-90 transition-all text-center">
                      {loading ? "Importing..." : "Upload CSV"}
                    </div>
                  </div>
                </div>
              </div>

              <div className="p-6 bg-brand-bg rounded-xl border border-brand-border">
                <div className="flex gap-4 items-start">
                  <AlertCircle className="w-5 h-5 text-brand-accent shrink-0 mt-0.5" />
                  <div className="space-y-1">
                    <p className="text-[11px] font-black uppercase tracking-widest text-brand-primary">Import Guidelines</p>
                    <ul className="text-[10px] text-brand-muted font-medium space-y-1 list-disc pl-4">
                      <li>Ensure that the "Organisation" column is populated for every row.</li>
                      <li>Standard lead statuses will be automatically mapped to your board columns.</li>
                      <li>Dates should be in YYYY-MM-DD or standard readable text format.</li>
                      <li>Maximum recommended batch size: 1,000 records per upload.</li>
                    </ul>
                  </div>
                </div>
              </div>

              <div className="bg-white border border-brand-border rounded-2xl p-8 space-y-6">
                <div>
                  <h4 className="text-sm font-black uppercase tracking-widest text-brand-primary mb-1">Duplicate Lead Management</h4>
                  <p className="text-[10px] text-brand-muted font-medium">Scan and resolve duplicate organisation records.</p>
                </div>

                <div className="flex justify-between items-center bg-brand-bg/50 p-4 rounded-xl border border-brand-border">
                  <div className="flex items-center gap-3">
                    <div className="p-2 bg-white rounded-lg border border-brand-border">
                      <Target className="w-4 h-4 text-brand-accent" />
                    </div>
                    <div>
                      <p className="text-[10px] font-black uppercase text-brand-primary">Scan Database</p>
                      <p className="text-[9px] text-brand-muted font-medium">Identify identical organisation names (case-insensitive)</p>
                    </div>
                  </div>
                  <button 
                    onClick={findDuplicates}
                    className="px-6 py-2 bg-brand-accent text-white rounded-lg text-[10px] font-black uppercase tracking-widest hover:bg-brand-accent/90 transition-all shadow-md"
                  >
                    Scan for Duplicates
                  </button>
                </div>

                {duplicateGroups.length > 0 && (
                  <div className="space-y-4 max-h-[500px] overflow-y-auto pr-2">
                    {duplicateGroups.map((group, gIdx) => (
                      <div key={gIdx} className="bg-white border border-brand-border rounded-xl overflow-hidden shadow-sm">
                        <div className="bg-brand-bg px-4 py-2 border-b border-brand-border flex justify-between items-center">
                          <span className="text-[9px] font-black uppercase text-brand-primary">{group[0].organisation}</span>
                          <span className="text-[8px] font-black text-rose-600 bg-rose-50 px-2 py-0.5 rounded-full border border-rose-100">{group.length} Duplicates</span>
                        </div>
                        <div className="divide-y divide-brand-border">
                          {group.map((lead: any) => (
                            <div key={lead.id} className="p-4 flex items-center justify-between hover:bg-brand-bg/5 transition-colors">
                              <div className="space-y-0.5">
                                <div className="flex items-center gap-2">
                                  <span className="text-[9px] font-bold text-brand-primary">{lead.leadStatus || "No Status"}</span>
                                  <span className="text-[9px] text-brand-muted truncate max-w-[150px]">{lead.email || "No Email"}</span>
                                </div>
                                <p className="text-[8px] font-mono text-brand-muted/50 uppercase">ID: {lead.id}</p>
                              </div>
                              <button 
                                onClick={() => deleteDuplicateLead(lead.id, lead.organisation)}
                                disabled={loading}
                                className={cn(
                                  "p-1.5 rounded-md transition-all",
                                  confirmDeleteId === lead.id ? "bg-rose-600 text-white" : "text-rose-500 hover:text-rose-700 hover:bg-rose-50",
                                  loading && "opacity-30"
                                )}
                                title={confirmDeleteId === lead.id ? "Confirm?" : "Delete this entry"}
                              >
                                {confirmDeleteId === lead.id ? <CheckCircle2 className="w-3.5 h-3.5 animate-pulse" /> : <Trash2 className="w-3.5 h-3.5" />}
                              </button>
                            </div>
                          ))}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
                {duplicateGroups.length === 0 && (
                  <div className="text-center py-6 opacity-30">
                    <ShieldCheck className="w-6 h-6 mx-auto mb-2" />
                    <p className="text-[9px] font-black uppercase tracking-widest">Database consistency verified</p>
                  </div>
                )}
              </div>
            </div>
          </motion.div>
        )}
        {activeAdminTab === "archive" && (
          <motion.div key="archive" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} className="space-y-8">
             <div className="bg-white border border-brand-border rounded-2xl p-8 card-shadow">
                <div className="flex justify-between items-center mb-6">
                  <div>
                    <h4 className="text-sm font-black text-brand-primary uppercase tracking-tight">Archived Leads</h4>
                    <p className="text-[10px] text-brand-muted font-medium">Deleted leads waiting for restoration or permanent removal</p>
                  </div>
                  <div className="flex items-center gap-2 text-[10px] font-bold text-brand-muted bg-brand-bg px-3 py-1.5 rounded-full border border-brand-border">
                    <Archive className="w-3" />
                    {archivedLeads.length} Records
                  </div>
                </div>

                {archivedLeads.length === 0 ? (
                  <div className="text-center py-20 opacity-30 border-2 border-dashed border-brand-border rounded-2xl">
                    <Archive className="w-12 h-12 mx-auto mb-4" />
                    <p className="text-[10px] font-black uppercase tracking-widest leading-relaxed">The archive is empty.<br/>Safe & Secure.</p>
                  </div>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full text-left">
                      <thead>
                        <tr className="border-b border-brand-border">
                          <th className="px-6 py-4 text-[10px] font-black text-brand-muted uppercase tracking-widest">Organisation</th>
                          <th className="px-6 py-4 text-[10px] font-black text-brand-muted uppercase tracking-widest">Archived Date</th>
                          <th className="px-6 py-4 text-[10px] font-black text-brand-muted uppercase tracking-widest">Deleted By</th>
                          <th className="px-6 py-4 text-right text-[10px] font-black text-brand-muted uppercase tracking-widest">Actions</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-brand-border">
                        {archivedLeads.map((lead) => (
                          <tr key={lead.id} className="hover:bg-brand-bg/20 transition-colors group">
                            <td className="px-6 py-4">
                              <p className="text-xs font-black text-brand-primary">{lead.organisation}</p>
                              <p className="text-[9px] text-brand-muted">{lead.email || "-"}</p>
                            </td>
                            <td className="px-6 py-4">
                              <p className="text-[10px] font-bold text-brand-muted">{formatDate(lead.archivedAt)}</p>
                            </td>
                            <td className="px-6 py-4">
                              <p className="text-[10px] font-bold text-brand-primary">{lead.archivedBy || "System"}</p>
                            </td>
                            <td className="px-6 py-4 text-right">
                              <div className="flex justify-end gap-2 pr-2 opacity-0 group-hover:opacity-100 transition-opacity">
                                <button 
                                  onClick={() => restoreLead(lead)}
                                  className="px-3 py-1.5 bg-brand-accent text-white rounded-lg text-[9px] font-black uppercase tracking-widest hover:scale-105 transition-all flex items-center gap-1.5 whitespace-nowrap"
                                >
                                  <RotateCcw className="w-3 h-3" />
                                  Restore
                                </button>
                                <button 
                                  onClick={() => permanentlyDeleteLead(lead.id)}
                                  className={cn(
                                    "p-1.5 rounded-lg transition-all border border-brand-border",
                                    confirmDeleteId === lead.id ? "bg-rose-600 text-white border-rose-600" : "text-rose-500 hover:bg-rose-50"
                                  )}
                                  title={confirmDeleteId === lead.id ? "Click again to confirm" : "Permanently Delete"}
                                >
                                  {confirmDeleteId === lead.id ? <CheckCircle2 className="w-3.5 h-3.5" /> : <Trash2 className="w-3.5 h-3.5" />}
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
             </div>
          </motion.div>
        )}

        {activeAdminTab === "permissions" && (
          <motion.div key="permissions" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} className="space-y-8">
            <div className="bg-white border border-brand-border rounded-2xl p-8 card-shadow shadow-sm">
              <div className="flex justify-between items-center mb-8">
                <div>
                  <h4 className="text-xs font-black text-brand-primary uppercase tracking-widest">Assessor Permissions</h4>
                  <p className="text-[10px] text-brand-muted font-medium mt-1">Configure what users with the 'Assessor' role can do in the platform.</p>
                </div>
                <button 
                  onClick={async () => {
                    setLoading(true);
                    try {
                      const defaultPerms = {
                        canAddLead: false,
                        canEditLead: true,
                        canDeleteLead: false,
                        disallowedStatuses: ["Assessed", "Approved for future project"],
                        canRunResearch: true
                      };
                      await setDoc(doc(db, "settings", "global"), {
                        assessorPermissions: settings?.assessorPermissions || defaultPerms
                      }, { merge: true });
                      notify("Assessor permissions saved", "success");
                    } catch (err) {
                      notify("Failed to update permissions", "error");
                    } finally {
                      setLoading(false);
                    }
                  }}
                  className="flex items-center gap-2 bg-brand-primary text-white px-6 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-brand-primary/90 transition-all shadow-lg"
                >
                  <Save className="w-3.5 h-3.5" /> Save Changes
                </button>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-10">
                <div className="space-y-6">
                  <h5 className="text-[10px] font-bold text-brand-muted uppercase tracking-widest border-b border-brand-border pb-2">Functional Rights</h5>
                  <div className="space-y-4">
                    {[
                      { id: 'canAddLead', label: 'Create New Leads', description: 'Allow assessors to manually add new non-profits.' },
                      { id: 'canEditLead', label: 'Edit Existing Leads', description: 'Allow updating basic info and contact details.' },
                      { id: 'canDeleteLead', label: 'Archive/Delete leads', description: 'Allow removing leads from the active pipeline.' },
                      { id: 'canRunResearch', label: 'Trigger AI Research', description: 'Allow running Gemini research tasks.' }
                    ].map(p => (
                      <div key={p.id} className="flex items-center justify-between group p-3 hover:bg-brand-bg rounded-xl transition-colors">
                        <div>
                          <p className="text-xs font-black text-brand-primary">{p.label}</p>
                          <p className="text-[9px] text-brand-muted font-medium">{p.description}</p>
                        </div>
                        <button
                          onClick={async () => {
                            const current = settings?.assessorPermissions || {
                              canAddLead: false,
                              canEditLead: true,
                              canDeleteLead: false,
                              disallowedStatuses: ["Assessed", "Approved for future project"],
                              canRunResearch: true
                            };
                            const newPerms = {
                              ...current,
                              [p.id]: !current[p.id]
                            };
                            await setDoc(doc(db, "settings", "global"), { assessorPermissions: newPerms }, { merge: true });
                          }}
                          className={cn(
                            "w-12 h-6 rounded-full transition-all relative",
                            settings?.assessorPermissions?.[p.id] ? "bg-emerald-500 shadow-[0_0_10px_rgba(16,185,129,0.3)]" : "bg-brand-border"
                          )}
                        >
                          <div className={cn(
                            "absolute top-1 w-4 h-4 rounded-full bg-white transition-all shadow-sm",
                            settings?.assessorPermissions?.[p.id] ? "left-7" : "left-1"
                          )} />
                        </button>
                      </div>
                    ))}
                  </div>
                </div>

                <div className="space-y-6">
                  <h5 className="text-[10px] font-bold text-brand-muted uppercase tracking-widest border-b border-brand-border pb-2">Restricted Lead Statuses</h5>
                  <div className="grid grid-cols-1 gap-2">
                    {(settings?.leadsConfig?.statuses || []).map((status: string) => {
                      const isDisallowed = (settings?.assessorPermissions?.disallowedStatuses || []).includes(status);
                      return (
                        <button
                          key={status}
                          onClick={async () => {
                            const current = settings?.assessorPermissions?.disallowedStatuses || [];
                            const newList = isDisallowed 
                              ? current.filter((s: string) => s !== status)
                              : [...current, status];
                            const newPerms = {
                              ...(settings?.assessorPermissions || {
                                canAddLead: false,
                                canEditLead: true,
                                canDeleteLead: false,
                                disallowedStatuses: ["Assessed", "Approved for future project"],
                                canRunResearch: true
                              }),
                              disallowedStatuses: newList
                            };
                            await setDoc(doc(db, "settings", "global"), { assessorPermissions: newPerms }, { merge: true });
                          }}
                          className={cn(
                            "flex items-center justify-between p-3 rounded-xl border text-left transition-all group",
                            isDisallowed 
                              ? "bg-rose-50 border-rose-200 text-rose-700" 
                              : "bg-white border-brand-border text-brand-muted hover:border-brand-accent hover:text-brand-primary"
                          )}
                        >
                          <span className="text-[10px] font-black uppercase tracking-widest">{status}</span>
                          {isDisallowed ? (
                            <div className="flex items-center gap-1.5 px-3 py-1 bg-rose-500 text-white rounded-lg text-[8px] font-black uppercase tracking-widest shadow-lg shadow-rose-500/20">
                              <Lock className="w-2.5 h-2.5" /> Restricted
                            </div>
                          ) : (
                            <div className="flex items-center gap-1.5 px-3 py-1 bg-emerald-500 text-white rounded-lg text-[8px] font-black uppercase tracking-widest opacity-0 group-hover:opacity-100 transition-opacity">
                              <ShieldCheck className="w-2.5 h-2.5" /> Allowed
                            </div>
                          )}
                        </button>
                      );
                    })}
                  </div>
                </div>
              </div>
            </div>
          </motion.div>
        )}


      </AnimatePresence>
    </motion.div>
  );
}

function NavItem({ active, onClick, icon: Icon, label }: { active: boolean, onClick: () => void, icon: any, label: string }) {
  return (
    <button 
      onClick={onClick}
      className={cn(
        "w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-all",
        active ? "bg-white/10 text-white shadow-lg" : "text-white/40 hover:text-white hover:bg-white/5"
      )}
    >
      <Icon className={cn("w-4 h-4", active ? "text-brand-accent" : "")} />
      <span className="text-[13px] font-semibold">{label}</span>
    </button>
  );
}

// --- Components ---

function SortableStatusItem({ 
  id, 
  boardStatuses, 
  setBoardStatuses,
  columnWidths,
  setColumnWidths,
  onUpdateGlobal,
  readOnly
}: { 
  id: string, 
  boardStatuses: string[], 
  setBoardStatuses: React.Dispatch<React.SetStateAction<string[]>>,
  columnWidths: Record<string, number>,
  setColumnWidths: React.Dispatch<React.SetStateAction<Record<string, number>>>,
  onUpdateGlobal?: (newList: string[]) => void,
  readOnly?: boolean
}) {
  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging
  } = useSortable({ id, disabled: readOnly });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
    zIndex: isDragging ? 50 : 1
  };

  const handleStatusEdit = (newValue: string) => {
    if (readOnly || !newValue || newValue === id) return;
    const newList = boardStatuses.map(s => s === id ? newValue : s);
    setBoardStatuses(newList);
    if (onUpdateGlobal) onUpdateGlobal(newList);
    if (columnWidths[id]) {
      const newWidths = { ...columnWidths };
      newWidths[newValue] = newWidths[id];
      delete newWidths[id];
      setColumnWidths(newWidths);
    }
  };

  const handleStatusDelete = () => {
    if (readOnly) return;
    const newList = boardStatuses.filter(s => s !== id);
    setBoardStatuses(newList);
    if (onUpdateGlobal) onUpdateGlobal(newList);
  };

  return (
    <div 
      ref={setNodeRef} 
      style={style} 
      className={cn(
        "flex flex-col gap-3 p-4 bg-brand-bg border border-brand-border rounded-xl group transition-all w-full overflow-hidden",
        isDragging ? "shadow-2xl scale-[1.02] border-brand-accent/50 bg-white z-50" : "hover:border-brand-accent/30",
        readOnly ? "opacity-90" : ""
      )}
    >
      <div className="flex items-center gap-3 min-w-0">
        {!readOnly && (
          <div 
            {...attributes} 
            {...listeners} 
            className="cursor-grab active:cursor-grabbing p-1.5 hover:bg-white rounded-lg transition-colors text-brand-muted hover:text-brand-accent shrink-0"
          >
            <GripVertical className="w-4 h-4" />
          </div>
        )}
        <input 
          className={cn(
            "text-[11px] font-bold text-brand-primary flex-1 bg-transparent border-none focus:ring-0 outline-none truncate min-w-0",
            readOnly ? "cursor-default" : ""
          )}
          value={id}
          readOnly={readOnly}
          placeholder="Status Name"
          onChange={(e) => handleStatusEdit(e.target.value)}
        />
        <div className="flex items-center gap-2 shrink-0">
          <div className="flex items-center gap-1 border border-brand-border rounded-lg bg-white px-2 py-1">
            <span className="text-[8px] font-black text-brand-muted uppercase tracking-tighter">W</span>
            <input 
              type="number"
              className="w-10 text-[10px] font-bold text-brand-primary border-none p-0 focus:ring-0 text-center"
              value={columnWidths[id] || 320}
              onChange={(e) => {
                const val = parseInt(e.target.value);
                setColumnWidths(prev => ({ ...prev, [id]: val }));
              }}
            />
          </div>
          {!readOnly && (
            <button 
              onClick={handleStatusDelete}
              className="p-1.5 text-brand-muted hover:text-red-500 hover:bg-red-50 rounded-md transition-colors opacity-0 group-hover:opacity-100"
            >
              <Trash2 className="w-3.5 h-3.5" />
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

function Dashboard({ assessments, leads = [], onSelectAssessment, onSelectResearch, onDeletePartner, userProfile }: { 
  assessments: any[], 
  leads: any[],
  onSelectAssessment: (id: string) => void,
  onSelectResearch: (id: string) => void,
  onDeletePartner: (id: string) => void,
  userProfile?: any
}) {
  const isAdmin = userProfile?.role?.toLowerCase() === "admin";
  const [confirmingPartnerId, setConfirmingPartnerId] = useState<string | null>(null);
  const filteredLeads = leads.filter(l => {
    const status = (l.leadStatus || "").toLowerCase();
    return status !== "completed" && status !== "project completed";
  });

  const activeAssessments = filteredLeads.filter(l => (l.leadStatus || "").toLowerCase() === "under assessment").length;
  const assessedCount = filteredLeads.filter(l => (l.leadStatus || "").toLowerCase() === "assessed").length;
  const approvedFutureCount = filteredLeads.filter(l => (l.leadStatus || "").toLowerCase() === "approved for future project").length;
  const paidEngagementsCount = filteredLeads.filter(l => l.isPaidEngagement).length;
  
  const handlePartnerClick = (npId: string) => {
    // Find the latest assessment for this partner
    const assessment = assessments.find(a => a.nonProfitId === npId);
    if (assessment) {
      onSelectAssessment(assessment.id);
    } else {
      // If no assessment exists, we might want to create one, but for now just navigate
      // Actually, let's just trigger the 'add' flow by setting id to something special or handling in AssessmentsView
      onSelectAssessment(`new:${npId}`);
    }
  };

  return (
    <motion.div 
      initial={{ opacity: 0, y: 10 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -10 }}
      className="max-w-6xl mx-auto"
    >
      <div className="flex justify-between items-center mb-8">
        <h3 className="text-[10px] font-black uppercase tracking-[0.2em] text-brand-muted">Performance Overview</h3>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 mb-10">
        <StatsCard icon={Activity} label="In Assessment" value={activeAssessments.toString()} subLabel="Active Pipeline" />
        <StatsCard icon={CheckCircle2} label="Assessed" value={assessedCount.toString()} subLabel="Ready for Review" />
        <StatsCard icon={Target} label="Future Projects" value={approvedFutureCount.toString()} subLabel="Approved" />
        <StatsCard icon={DollarSign} label="Paid Leads" value={paidEngagementsCount.toString()} subLabel="Engagements" />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        <div className="bg-white p-8 rounded-xl border border-brand-border card-shadow">
          {filteredLeads.filter(l => (l.leadStatus || "").toLowerCase() === "under assessment").length > 0 && (
            <div className="mb-10 pb-10 border-b border-brand-border/50">
              <h3 className="text-[10px] font-bold uppercase tracking-widest text-brand-muted mb-6 flex items-center gap-2">
                <Activity className="w-3 h-3 text-amber-500" /> ASSESSMENT PIPELINE
              </h3>
              <div className="space-y-2">
                {filteredLeads.filter(l => (l.leadStatus || "").toLowerCase() === "under assessment").map(lead => (
                  <div 
                    key={lead.id}
                    className="w-full flex items-center justify-between p-4 rounded-lg bg-amber-50/30 border border-amber-100 hover:border-amber-200 transition-all group cursor-pointer"
                    onClick={() => onSelectResearch(lead.id)}
                  >
                    <div className="flex items-center gap-4">
                      <div className="w-2 h-2 rounded-full bg-amber-400 animate-pulse" />
                      <div>
                        <p className="font-bold text-sm text-brand-primary group-hover:text-brand-accent transition-colors">{lead.organisation}</p>
                        {lead.needsFurtherAssessment ? (
                          <div className="flex items-center gap-1.5 mt-0.5">
                            <AlertCircle className="w-3 h-3 text-amber-500" />
                            <p className="text-[10px] text-amber-600 font-bold uppercase tracking-widest">Further Assessment Required</p>
                          </div>
                        ) : (
                          <p className="text-[10px] text-brand-muted font-medium italic truncate max-w-[200px]">{lead.website || "Discovery lead"}</p>
                        )}
                      </div>
                    </div>
                    <div className="flex items-center gap-4">
                      <span className="text-[8px] font-black uppercase tracking-widest px-2 py-1 rounded bg-amber-100 text-amber-700 border border-amber-200">
                        Under Assessment
                      </span>
                      <ChevronRight className="w-4 h-4 text-brand-muted translate-x-[-10px] group-hover:translate-x-0 transition-transform" />
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {filteredLeads.filter(l => l.leadStatus === "Assessed").length > 0 && (
            <div className="mb-10 pb-10 border-b border-brand-border/50">
              <h3 className="text-[10px] font-bold uppercase tracking-widest text-brand-muted mb-6 flex items-center gap-2">
                <CheckCircle2 className="w-3 h-3 text-emerald-500" /> RECENTLY ASSESSED
              </h3>
              <div className="space-y-2">
                {filteredLeads.filter(l => l.leadStatus === "Assessed").map(lead => (
                  <div 
                    key={lead.id}
                    className="w-full flex items-center justify-between p-4 rounded-lg bg-emerald-50/30 border border-emerald-100 hover:border-emerald-200 transition-all group cursor-pointer"
                    onClick={() => onSelectResearch(lead.id)}
                  >
                    <div className="flex items-center gap-4">
                      <div className="w-2 h-2 rounded-full bg-emerald-400" />
                      <div>
                        <p className="font-bold text-sm text-brand-primary group-hover:text-brand-accent transition-colors">{lead.organisation}</p>
                          {lead.needsFurtherAssessment ? (
                            <div className="flex items-center gap-1.5 mt-0.5">
                              <AlertCircle className="w-3 h-3 text-amber-500" />
                              <p className="text-[10px] text-amber-600 font-bold uppercase tracking-widest">Further Assessment Required</p>
                            </div>
                          ) : lead.leadStatus === "Assessed" && !lead.approved ? (
                            <div className="flex items-center gap-1.5 mt-0.5">
                              <Clock className="w-3 h-3 text-blue-500" />
                              <p className="text-[10px] text-blue-600 font-bold uppercase tracking-widest">Waiting for Approval from DCM Board</p>
                            </div>
                          ) : lead.assessedBy ? (
                            <p className="text-[10px] text-brand-muted font-medium italic truncate max-w-[200px]">Assessed by {lead.assessedBy.split('@')[0] || "Assessor"}</p>
                          ) : null}
                      </div>
                    </div>
                    <div className="flex items-center gap-4">
                       {lead.needsFurtherAssessment ? (
                         <span className="text-[8px] font-black uppercase tracking-widest px-2 py-1 rounded bg-amber-100 text-amber-700 border border-amber-200">
                          Re-Assessment
                        </span>
                       ) : (
                         <span className="text-[8px] font-black uppercase tracking-widest px-2 py-1 rounded bg-emerald-100 text-emerald-700 border border-emerald-200">
                          Assessed
                        </span>
                       )}
                      <ChevronRight className="w-4 h-4 text-brand-muted translate-x-[-10px] group-hover:translate-x-0 transition-transform" />
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          <h3 className="text-[10px] font-bold uppercase tracking-widest text-brand-muted mb-6">PIPELINE HISTORY</h3>
          <div className="space-y-2">
            {filteredLeads
              .filter(l => assessments.some(a => a.nonProfitId === l.id))
              .sort((a, b) => {
                const assA = assessments.find(ass => ass.nonProfitId === a.id);
                const assB = assessments.find(ass => ass.nonProfitId === b.id);
                return new Date(assB.updatedAt || assB.createdAt).getTime() - new Date(assA.updatedAt || assA.createdAt).getTime();
              })
              .slice(0, 15)
              .map(lead => {
                const assessment = assessments.find(a => a.nonProfitId === lead.id);
                
                // Determine display status based on lead status (primary) and assessment status (secondary)
                let statusLabel = "Pending";
                let dotColor = "bg-amber-500";
                let badgeStyle = "bg-amber-50 text-amber-600 border-amber-200";

                if (lead.approved || lead.leadStatus === "Approved for Future Project") {
                  statusLabel = "Approved";
                  dotColor = "bg-emerald-500";
                  badgeStyle = "bg-emerald-50 text-emerald-600 border-emerald-200";
                } else if ((lead.leadStatus || "").toLowerCase().includes("under assessment")) {
                  statusLabel = "Under assessment";
                  dotColor = "bg-blue-500";
                  badgeStyle = "bg-blue-50 text-blue-600 border-blue-200";
                } else if ((lead.leadStatus || "").toLowerCase() === "assessed") {
                  statusLabel = "Assessed";
                  dotColor = "bg-emerald-500";
                  badgeStyle = "bg-emerald-50 text-emerald-600 border-emerald-200";
                } else if (lead.leadStatus === "Not Suitable" || assessment?.status === "not suitable") {
                  statusLabel = "Not Suitable";
                  dotColor = "bg-rose-500";
                  badgeStyle = "bg-rose-50 text-rose-600 border-rose-200";
                } else if (lead.leadStatus === "Completed" || lead.leadStatus === "Project Completed") {
                  statusLabel = "Completed";
                  dotColor = "bg-brand-primary";
                  badgeStyle = "bg-brand-bg text-brand-primary border-brand-border";
                } else if (assessment?.status === "approved") {
                  statusLabel = "Approved";
                  dotColor = "bg-emerald-500";
                  badgeStyle = "bg-emerald-50 text-emerald-600 border-emerald-200";
                }
                
                return (
                  <div 
                    key={lead.id} 
                    onClick={() => handlePartnerClick(lead.id)}
                    className="w-full flex items-center justify-between p-4 rounded-lg border border-transparent hover:border-brand-border hover:bg-brand-bg transition-all group cursor-pointer"
                  >
                  <div className="flex items-center gap-4">
                    <div className={cn("w-2 h-2 rounded-full", dotColor)} />
                    <div>
                      <p className="font-bold text-sm text-brand-primary group-hover:text-brand-accent transition-colors">{lead.organisation}</p>
                      <p className="text-[10px] text-brand-muted font-medium italic truncate max-w-[200px]">{lead.website || "No website listed"}</p>
                    </div>
                  </div>
                  <div className="flex items-center gap-4">
                    <span className={cn(
                      "hidden md:block text-[8px] font-black uppercase tracking-widest px-2 py-1 rounded border",
                      badgeStyle
                    )}>
                      {statusLabel}
                    </span>
                    <div className="flex items-center gap-2 opacity-0 group-hover:opacity-100 transition-all">
                      {isAdmin && (
                        confirmingPartnerId === lead.id ? (
                          <div className="flex items-center gap-1 bg-red-50 p-1 rounded-lg border border-red-100">
                            <button 
                              onClick={(e) => { 
                                e.stopPropagation(); 
                                onDeletePartner(lead.id);
                                setConfirmingPartnerId(null);
                              }}
                              className="text-[10px] font-bold text-red-600 px-2 py-1 hover:bg-red-100 rounded transition-all"
                            >
                              Confirm Delete?
                            </button>
                            <button 
                              onClick={(e) => { 
                                e.stopPropagation(); 
                                setConfirmingPartnerId(null);
                              }}
                              className="p-1 text-slate-400 hover:text-slate-600 rounded"
                            >
                              <X className="w-3 h-3" />
                            </button>
                          </div>
                        ) : (
                          <button 
                            onClick={(e) => { 
                              e.stopPropagation(); 
                              setConfirmingPartnerId(lead.id); 
                            }}
                            className="p-2 text-brand-muted hover:text-red-500 hover:bg-red-50 rounded-lg transition-all"
                            title="Remove Partner"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        )
                      )}
                      <ChevronRight className="w-4 h-4 text-brand-muted translate-x-[-10px] group-hover:translate-x-0 transition-transform" />
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        <div className="bg-white p-8 rounded-xl border border-brand-border card-shadow flex flex-col">
          <h3 className="text-[10px] font-bold uppercase tracking-widest text-brand-muted mb-6">Pipeline Distribution</h3>
          <div className="flex-1 min-h-[250px]">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart margin={{ top: 10, right: 80, bottom: 10, left: 80 }}>                <Pie 
                  data={[
                    { name: 'Under assessment', value: filteredLeads.filter(l => (l.leadStatus || "").toLowerCase().includes("under assessment")).length },
                    { name: 'Assessed', value: filteredLeads.filter(l => (l.leadStatus || "").toLowerCase() === "assessed").length },
                    { name: 'Approved Project', value: filteredLeads.filter(l => (l.leadStatus || "").toLowerCase().includes("approved for future")).length }
                  ]} 
                  innerRadius={55} 
                  outerRadius={80} 
                  paddingAngle={5} 
                  dataKey="value"
                  stroke="none"
                  label={({ name, cx, cy, midAngle, innerRadius, outerRadius, value }) => {
                    if (value === 0) return null;
                    const RADIAN = Math.PI / 180;
                    const radius = outerRadius + 25;
                    const x = cx + radius * Math.cos(-midAngle * RADIAN);
                    const y = cy + radius * Math.sin(-midAngle * RADIAN);
                    return (
                      <text 
                        x={x} 
                        y={y} 
                        fill="#64748b" 
                        textAnchor={x > cx ? 'start' : 'end'} 
                        dominantBaseline="central"
                        className="text-[7.5px] font-black uppercase tracking-tight"
                      >
                        {name.split(' ')[0]} ({value})
                      </text>
                    );
                  }}
                >
                  <Cell fill="#F59E0B" />
                  <Cell fill="#10B981" />
                  <Cell fill="#06B6D4" />
                </Pie>
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div className="flex flex-wrap justify-center gap-4 mt-4">
            <LegendItem label="Under assessment" color="#F59E0B" />
            <LegendItem label="Assessed" color="#10B981" />
            <LegendItem label="Approved" color="#06B6D4" />
          </div>
        </div>
      </div>
    </motion.div>
  );
}

function ResearchTool({ leads, userProfile, initialPartnerId, onPartnerSelect, onDeletePartner, settings, aiConfig, onShowAiConfig, handleAiError, notify, logLeadEvent }: { 
  leads?: any[], 
  userProfile: any,
  initialPartnerId?: string | null,
  onPartnerSelect: (id: string) => void,
  onDeletePartner: (id: string) => void,
  settings: any,
  aiConfig: any,
  onShowAiConfig: () => void,
  handleAiError: (err: any, source: string) => void,
  notify: (msg: string, type: 'success' | 'error') => void,
  logLeadEvent: (leadId: string, type: string, description: string, metadata?: any) => Promise<void>
}) {
  const [confirmingDelete, setConfirmingDelete] = useState(false);
  const [query, setQuery] = useState("");
  const [linkedinUrl, setLinkedinUrl] = useState("");
  const [loading, setLoading] = useState(false);
  const [linkedinLoading, setLinkedinLoading] = useState(false);
  const [result, setResult] = useState<NonProfitResearch | null>(null);
  const [saving, setSaving] = useState(false);
  const [selectedPartnerId, setSelectedPartnerId] = useState("");

  useEffect(() => {
    if (initialPartnerId && leads && leads.length > 0) {
      const partner = leads.find(l => l.id === initialPartnerId);
      if (partner) {
        // If we don't have a result yet, or if the ID changed, load it
        // We also want to sync if the lead data changed in the background
        if (selectedPartnerId !== initialPartnerId || !result) {
          handlePartnerSelect(initialPartnerId);
        }
      }
    }
  }, [initialPartnerId, leads]); 

  const handleSearch = async (e?: React.FormEvent, existingQuery?: string) => {
    if (e) e.preventDefault();
    const q = (existingQuery || query || "").trim();
    if (!q) return;
    
    // First priority: use current selectedPartnerId if it matches the current research
    let partner = leads?.find(l => l.id === selectedPartnerId);
    
    // Second priority: use initialPartnerId if provided and no partner is selected yet
    if (!partner && initialPartnerId) {
      partner = leads?.find(l => l.id === initialPartnerId);
    }
    
    // Third priority: search by name
    if (!partner) {
      partner = leads?.find(l => {
        const orgName = (l.organisation || "").toLowerCase().trim();
        const leadName = (l.name || "").toLowerCase().trim();
        const queryLower = q.toLowerCase().trim();
        
        return orgName === queryLower || 
               leadName === queryLower ||
               (orgName.length > 3 && queryLower.includes(orgName)) ||
               (queryLower.length > 3 && orgName.includes(queryLower));
      });
    }
    
    setLoading(true);
    setResult(null);
    // Don't clear selectedPartnerId if we found a partner - keep the context!
    if (!partner) {
      setSelectedPartnerId("");
    } else {
      setSelectedPartnerId(partner.id);
      onPartnerSelect(partner.id);
    }

    try {
      const data = await researchNonProfit(q, partner?.website, aiConfig);
      setResult(data);
      notify("Organization intelligence retrieved", "success");
      
      // Auto-save back to DCM Leads record if we have a partner
      if (partner) {
        const docRef = doc(db, "leads", partner.id);
        
        // Smart merge: Only include fields that have meaningful values from AI
        const updates: any = {
          updatedAt: new Date().toISOString()
        };
        
        // If the record has no mission/activity, use AI name/mission
        if (!partner.organisation) updates.organisation = data.name;
        if (!partner.activity) updates.activity = data.mission || data.products_services;

        // Fields from AI to merge if NOT empty
        const fieldsToSync = [
          'revenue', 'ein', 'website', 'mission', 'products_services', 
          'staff_members', 'linkedin_url', 'linkedin_overview', 
          'staff_linkedin_summary', 'linkedin_activity',
          'charity_navigator_rating', 'propublica_grants'
        ];

        fieldsToSync.forEach(field => {
          const aiValue = (data as any)[field];
          const lowerAiVal = String(aiValue || "").toLowerCase();
          const isNotAvailable = !aiValue ||
                                lowerAiVal === "" || 
                                lowerAiVal.includes("no public record found") || 
                                lowerAiVal.includes("not found") || 
                                lowerAiVal.includes("n/a");

          if (aiValue && !isNotAvailable) {
            updates[field] = aiValue;
          }
        });

        await setDoc(docRef, updates, { merge: true });
        // After save, the 'leads' snapshot will trigger and we might want to refresh 'result'
        // but since 'data' is already the latest research result, we're good.
      }
    } catch (err: any) {
      handleAiError(err, "Non-Profit Research");
    } finally {
      setLoading(false);
    }
  };

  const handleLinkedInAudit = async () => {
    if (!linkedinUrl.trim() || !result) return;
    setLinkedinLoading(true);
    try {
      const data = await researchLinkedIn(linkedinUrl, result.name, aiConfig);
      const updatedResult = {
        ...result,
        linkedin_url: linkedinUrl,
        linkedin_overview: data.overview,
        staff_linkedin_summary: data.topPeople,
        linkedin_activity: data.topPosts
      };
      setResult(updatedResult);
      
      // Auto-save LinkedIn audit if we have a partner ID
      if (selectedPartnerId) {
        const docRef = doc(db, "leads", selectedPartnerId);
        await setDoc(docRef, updatedResult, { merge: true });
        notify("LinkedIn audit success & saved", "success");
      } else {
        notify("LinkedIn audit success (not saved - no record)", "success");
      }
    } catch (err: any) {
      handleAiError(err, "LinkedIn Audit");
    } finally {
      setLinkedinLoading(false);
    }
  };

  const handlePartnerSelect = (id: string) => {
    setSelectedPartnerId(id);
    onPartnerSelect(id);
    if (!id) {
      setResult(null);
      return;
    }
    const partner = leads?.find(l => l.id === id);
    if (partner) {
      setQuery(partner.organisation);
      setResult({
        name: partner.organisation || partner.name || "",
        revenue: partner.revenue || "",
        products_services: partner.products_services || partner.services || partner.activity || "",
        staff_members: partner.staff_members || partner.staffCount || "",
        mission: partner.mission || partner.activity || "",
        website: partner.website || "",
        linkedin_activity: partner.linkedin_activity || "",
        staff_linkedin_summary: partner.staff_linkedin_summary || "",
        linkedin_url: partner.linkedin_url || "",
        linkedin_overview: partner.linkedin_overview || "",
        ein: partner.ein || "",
        propublica_grants: partner.propublica_grants || "",
        charity_navigator_rating: partner.charity_navigator_rating || ""
      });
      setLinkedinUrl(partner.linkedin_url || "");
    }
  };

  const handleSaveToDb = async () => {
    if (!result) return;
    setSaving(true);
    try {
      console.log("Saving lead from Research to DB...", result);
      const lead = leads.find(l => l.id === selectedPartnerId);
      
      const data: any = {
        updatedAt: new Date().toISOString()
      };

      // Only update name/activity if we have actual content
      const aiName = result.name;
      if (aiName && !String(aiName).toLowerCase().includes("not found") && aiName !== "") {
        data.organisation = aiName;
      }

      const aiActivity = result.mission || result.products_services;
      if (aiActivity && !String(aiActivity).toLowerCase().includes("not found") && aiActivity !== "") {
        data.activity = aiActivity;
      }

      // Only copy result fields if they have content to avoid overwriting existing data with blanks
      Object.keys(result).forEach(key => {
        const val = (result as any)[key];
        const lowerVal = String(val || "").toLowerCase();
        const isNotAvailable = !val ||
                              lowerVal === "" || 
                              lowerVal.includes("no public record found") || 
                              lowerVal.includes("not found") || 
                              lowerVal.includes("n/a");
                              
        if (val && !isNotAvailable) {
          data[key] = val;
        }
      });

      if (selectedPartnerId && lead) {
        await setDoc(doc(db, "leads", selectedPartnerId), data, { merge: true });
        await logLeadEvent(selectedPartnerId, "RESEARCH_UPDATE", "Organization details updated via Deep Research");
        notify("Partner record updated in pipeline", "success");
      } else {
        // Clean ID if it somehow exists
        delete data.id;
        const docRef = await addDoc(collection(db, "leads"), {
          ...data,
          leadStatus: "First meeting/contact",
          createdAt: new Date().toISOString()
        });
        await logLeadEvent(docRef.id, "CREATED", "Lead created via Deep Research Discovery");
        notify("New lead added to pipeline", "success");
      }
    } catch (err: any) {
      console.error("Save to DB failed:", err);
      notify("Failed to save to database. Check permissions.", "error");
      handleFirestoreError(err, selectedPartnerId ? OperationType.UPDATE : OperationType.CREATE, "leads");
    } finally {
      setSaving(false);
    }
  };

  const deletePartner = async () => {
    if (!selectedPartnerId) return;
    try {
      await onDeletePartner(selectedPartnerId);
      setSelectedPartnerId("");
      setResult(null);
    } catch (err) {
      handleFirestoreError(err, OperationType.DELETE, "leads");
    }
  };

  const revenueInt = parseRevenueValue(result?.revenue || "");
  const currentAdminFee = revenueInt > 1000000 ? 1000 : 500;

  return (
    <motion.div 
      initial={{ opacity: 0, x: 20 }}
      animate={{ opacity: 1, x: 0 }}
      className="max-w-4xl mx-auto"
    >
      <div className="mb-12">
        <h2 className="text-4xl font-bold tracking-tighter mb-4 text-brand-primary">Deep Research Agent</h2>
        <p className="text-lg font-medium text-brand-muted">Explore non-profits or view existing partner research records.</p>
      </div>

      <div className="flex flex-col md:flex-row gap-4 mb-10">
        <form onSubmit={handleSearch} className="flex-1 relative group">
          <input 
            type="text" 
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search new organization name..."
            className="w-full bg-white border border-brand-border rounded-xl py-5 pl-8 pr-40 text-lg font-semibold tracking-tight outline-none focus:ring-2 focus:ring-brand-accent/20 transition-all card-shadow"
          />
          <button 
            disabled={loading}
            className="absolute right-3 top-3 bottom-3 px-8 bg-brand-accent text-white rounded-lg flex items-center justify-center gap-2 hover:bg-brand-accent/90 disabled:opacity-50 shadow-lg shadow-brand-accent/20 transition-all font-bold text-xs uppercase tracking-widest"
          >
            {loading ? <Loader2 className="w-5 h-5 animate-spin" /> : <SearchIcon className="w-5 h-5" />}
            RESEARCH
          </button>
        </form>

        <div className="md:w-1/4 bg-white border border-brand-border rounded-xl p-3 card-shadow flex items-center gap-3">
          <select 
            value={selectedPartnerId}
            onChange={(e) => handlePartnerSelect(e.target.value)}
            className="w-full bg-transparent border-none text-[10px] font-black uppercase tracking-widest outline-none cursor-pointer text-brand-primary"
          >
            <option value="">Under assessment leads</option>
            {leads && leads.filter(l => (l.leadStatus || "").toLowerCase().includes("under assessment")).length > 0 && (
              <optgroup label="DCM Leads">
                {leads.filter(l => (l.leadStatus || "").toLowerCase().includes("under assessment")).map(l => (
                  <option key={l.id} value={l.id}>{l.organisation} {l.isVerified ? '✓' : ''}</option>
                ))}
              </optgroup>
            )}
          </select>
        </div>
      </div>

      {query && leads && !result?.website && !leads.find(l => (l.organisation || l.name)?.toLowerCase() === query?.toLowerCase())?.website && (
        <div className="mb-8 flex items-center gap-4 px-8 py-4 bg-rose-50 border-2 border-rose-100 rounded-2xl text-xs font-bold text-rose-600 shadow-sm animate-pulse">
          <AlertCircle className="w-6 h-6 shrink-0" />
          <div className="flex-1">
            <p className="uppercase tracking-[0.1em] font-black leading-none mb-1">Missing Website Data</p>
            <p className="opacity-80">This organization is missing a URL in the Leads record. AI will attempt to discover it during discovery.</p>
          </div>
        </div>
      )}

      {result && (
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white border border-brand-border rounded-2xl p-10 card-shadow space-y-10"
        >
          <div className="flex flex-col md:flex-row justify-between items-start gap-6 pb-8 border-b border-brand-border">
              <div className="flex gap-6 items-center">
                <div className="w-16 h-16 bg-blue-50 rounded-2xl flex items-center justify-center font-bold text-brand-accent text-2xl shadow-sm border border-blue-100 shrink-0">
                  {(String(result.name || "NP")).substring(0, 2).toUpperCase()}
                </div>
                <div>
                  <h3 className="text-3xl font-extrabold tracking-tight text-brand-primary">{result.name || "Researching..."}</h3>
                  <div className="flex flex-wrap items-center gap-x-4 gap-y-1">
                    <a href={result.website} target="_blank" rel="noreferrer" className="text-sm font-medium text-brand-accent flex items-center gap-1 hover:underline">
                      <ExternalLink className="w-3 h-3" /> {result.website}
                    </a>
                    {result.ein && (
                      <span className="text-[10px] bg-brand-bg px-2 py-0.5 rounded border border-brand-border font-mono font-bold text-brand-muted">
                        EIN: {result.ein}
                      </span>
                    )}
                  </div>
                </div>
              </div>
              <div className="flex gap-3 w-full md:w-auto">
              {selectedPartnerId && (
                <>
                  <button 
                    onClick={() => handleSearch(undefined, result.name)}
                    disabled={loading}
                    className="flex-1 md:flex-none px-6 py-3 bg-brand-bg border border-brand-border rounded-lg font-bold text-xs uppercase tracking-widest text-brand-muted hover:text-brand-primary hover:bg-brand-border transition-all disabled:opacity-50 flex items-center justify-center gap-2"
                  >
                    {loading ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <SearchIcon className="w-3.5 h-3.5" />}
                    RE-RUN AI
                  </button>
                  {confirmingDelete ? (
                    <div className="flex items-center gap-2 bg-red-50 p-2 rounded-lg border border-red-100">
                      <span className="text-[10px] font-bold text-red-600 uppercase">Confirm Delete?</span>
                      <button 
                         onClick={() => {
                           deletePartner();
                           setConfirmingDelete(false);
                         }}
                         className="px-3 py-1 bg-red-600 text-white rounded text-[10px] font-black hover:bg-red-700"
                      >
                        YES
                      </button>
                      <button 
                         onClick={() => setConfirmingDelete(false)}
                         className="px-3 py-1 bg-white border border-slate-200 text-slate-600 rounded text-[10px] font-black hover:bg-slate-50"
                      >
                        NO
                      </button>
                    </div>
                  ) : (
                    (userProfile?.role?.toLowerCase() === "admin" || userProfile?.role?.toLowerCase() === "dcm boards") && (
                      <div className="flex gap-2">
                        <button 
                          onClick={async () => {
                            if (!selectedPartnerId) return;
                            const confirmed = window.confirm("CRITICAL: This will PERMANENTLY WIPE all AI-generated research, ProPublica data, and Charity Navigator ratings for this lead. You will need to RE-RUN AI to restore them. Continue?");
                            if (!confirmed) return;
                            try {
                              await setDoc(doc(db, "leads", selectedPartnerId), {
                                briefSummary: "",
                                revenue: "",
                                ein: "",
                                charity_navigator_rating: "",
                                propublica_grants: "",
                                linkedin_overview: "",
                                staff_linkedin_summary: "",
                                staff_members: "",
                                dcmComment: "",
                                updatedAt: new Date().toISOString()
                              }, { merge: true });
                              setResult(null);
                              setQuery("");
                              notify("Master Research Data Purged.", "success");
                            } catch (err) {
                              notify("Wipe failed", "error");
                            }
                          }}
                          className="p-3 bg-white border border-rose-200 text-rose-500 rounded-lg hover:bg-rose-50 transition-all"
                          title="Purge AI Research (Wipe Hallucinations)"
                        >
                          <RotateCw className="w-5 h-5" />
                        </button>
                        <button 
                          onClick={() => setConfirmingDelete(true)}
                          className="p-3 bg-white border border-red-200 text-red-500 rounded-lg hover:bg-red-50 transition-all group"
                          title="Delete Organization"
                        >
                          <Trash2 className="w-5 h-5 group-hover:scale-110 transition-transform" />
                        </button>
                      </div>
                    )
                  )}
                </>
              )}
              <button 
                onClick={handleSaveToDb}
                disabled={saving}
                className="flex-1 md:flex-none px-6 py-3 bg-brand-primary text-white border border-brand-primary rounded-lg font-bold text-xs uppercase tracking-widest hover:opacity-90 transition-all disabled:opacity-50"
              >
                {saving ? "SAVING..." : selectedPartnerId ? "UPDATE RECORD" : "ADD TO PIPELINE"}
              </button>
            </div>
          </div>

          {/* LinkedIn Audit Box */}
          <div className="bg-brand-bg rounded-2xl p-6 border border-brand-border">
            <div className="flex items-center gap-2 mb-4">
              <ShieldCheck className="w-4 h-4 text-brand-accent" />
              <h4 className="text-[10px] font-black uppercase tracking-wider text-brand-muted">LinkedIn Presence Audit</h4>
            </div>
            <div className="flex gap-4">
              <input 
                type="text" 
                value={linkedinUrl}
                onChange={(e) => setLinkedinUrl(e.target.value)}
                placeholder="Linkedin Org URL (e.g. linkedin.com/company/name)..."
                className="flex-1 bg-white border border-brand-border rounded-xl py-4 px-6 text-sm font-medium outline-none focus:ring-2 focus:ring-brand-accent/20"
              />
              <button 
                onClick={handleLinkedInAudit}
                disabled={linkedinLoading || !linkedinUrl}
                className="px-8 bg-brand-accent text-white rounded-xl font-bold text-xs uppercase tracking-widest hover:bg-brand-accent/90 disabled:opacity-50 flex items-center gap-2 whitespace-nowrap"
              >
                {linkedinLoading ? <Loader2 className="w-4 h-4 animate-spin" /> : <SearchIcon className="w-4 h-4" />}
                RUN AUDIT
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-10">
                <div className="space-y-10">
                  <div className="space-y-6">
                    <div>
                      <SectionHeader label="Organizational Insights" />
                      <div className="bg-brand-bg p-6 rounded-2xl border border-brand-border shadow-sm">
                        <p className="text-xs font-medium text-brand-primary leading-relaxed markdown-content tabular-nums">{result.revenue}</p>
                        <p className="text-[10px] font-bold text-brand-muted uppercase tracking-wider mt-3">Annual Revenue Footprint</p>
                      </div>
                    </div>
                    
                    <div>
                      <SectionHeader label="Mission & Impact" />
                      <div className="text-sm font-medium leading-relaxed text-brand-text/80 bg-brand-bg/30 p-6 rounded-2xl border border-brand-border/50 italic markdown-content shadow-inner">
                        <Markdown>{String(result.mission || "")}</Markdown>
                      </div>
                    </div>

                    <div>
                      <SectionHeader label="Grant History & Financials (ProPublica)" />
                      <div className="bg-brand-bg/50 p-6 rounded-2xl border border-brand-border text-xs font-medium leading-relaxed markdown-content">
                        <Markdown>{String(result.propublica_grants || "No grant data found. Run a new search or verify EIN.")}</Markdown>
                      </div>
                    </div>

                    <div>
                      <SectionHeader label="Charity Navigator Status" />
                      <div className="bg-brand-bg/50 p-6 rounded-2xl border border-brand-border text-xs font-medium leading-relaxed markdown-content">
                        <Markdown>{String(result.charity_navigator_rating || "No Rating available.")}</Markdown>
                      </div>
                    </div>
                  </div>

                  {result.linkedin_overview && (
                    <div>
                      <SectionHeader label="LinkedIn Overview" />
                      <div className="bg-brand-bg/50 p-6 rounded-2xl border border-brand-border text-xs font-medium leading-relaxed markdown-content">
                        <Markdown>{String(result.linkedin_overview)}</Markdown>
                      </div>
                    </div>
                  )}
                </div>

                <div className="space-y-10">
                  <div>
                    <SectionHeader label="Leadership & Key Staff" />
                    <div className="bg-brand-bg p-6 rounded-2xl border border-brand-border text-[13px] font-medium leading-relaxed text-brand-text markdown-content">
                       <Markdown>{String(result.staff_linkedin_summary || result.staff_members || "")}</Markdown>
                    </div>
                  </div>

                  <div>
                    <SectionHeader label="Engagement & Recent Posts" />
                    <div className="bg-brand-bg/30 p-6 rounded-2xl border border-brand-border/50 text-[12px] font-medium leading-relaxed text-brand-text markdown-content">
                       <Markdown>{String(result.linkedin_activity || "Perform a LinkedIn Audit to fetch recent engagement themes.")}</Markdown>
                    </div>
                  </div>

                  <div>
                    <SectionHeader label="Products & Services" />
                    <div className="text-[13px] font-medium leading-relaxed text-brand-text markdown-content bg-brand-bg/10 p-4 rounded-xl">
                       <Markdown>{String(result.products_services || "")}</Markdown>
                    </div>
                  </div>

                  <div className="p-6 bg-brand-primary text-white rounded-2xl relative overflow-hidden">
                    <div className="absolute top-0 right-0 p-6 opacity-10 pointer-events-none">
                      <DollarSign className="w-16 h-16" />
                    </div>
                    <p className="text-[10px] font-bold uppercase tracking-widest text-brand-accent mb-2">Funding Eligibility</p>
                    <p className="text-xs leading-relaxed text-white/80 font-medium">
                      Based on revenue indicators, this organization qualifies for a <strong className="text-white">${currentAdminFee.toLocaleString()} Project Admin Fee</strong> tier with standard reporting cycles.
                    </p>
                  </div>
                </div>
              </div>
            </motion.div>
          )}

          {!result && !loading && (
            <div className="p-32 text-center border-2 border-dashed border-brand-border rounded-3xl bg-white/50">
               <div className="w-20 h-20 bg-brand-bg rounded-3xl flex items-center justify-center mx-auto mb-8 text-brand-muted">
                  <SearchIcon className="w-8 h-8" />
               </div>
               <h4 className="text-xl font-black text-brand-primary uppercase tracking-tighter mb-2">Awaiting Intelligence Command</h4>
               <p className="text-brand-muted font-bold text-xs uppercase tracking-widest">Select a partner or search to initialize deep research.</p>
            </div>
          )}
      </motion.div>
    );
}

function AssessmentsView({ assessments, leads, selectedId, onClearSelection, settings, userProfile, aiConfig, onShowAiConfig, handleAiError, notify, onSetView }: { 
  assessments: any[], 
  leads: any[],
  selectedId: string | null, 
  onClearSelection: () => void,
  settings: any,
  userProfile: any,
  aiConfig: any,
  onShowAiConfig: () => void,
  handleAiError: (err: any, source: string) => void,
  notify: (msg: string, type: 'success' | 'error') => void,
  onSetView: (v: View, leadId?: string) => void
}) {
  const [confirmingAssessmentId, setConfirmingAssessmentId] = useState<string | null>(null);
  const [selectedNp, setSelectedNp] = useState<string>("");
  const isBoardsOrAdmin = userProfile?.role?.toLowerCase() === "admin" || userProfile?.role?.toLowerCase() === "dcm boards";
  const [scoring, setScoring] = useState(false);
  const [isNew, setIsNew] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);

  // Derive queue items exclusively from DCM Leads with "Under assessment" status
  const assessmentQueueItems = leads
    .filter(l => (l.leadStatus || "").toLowerCase() === "under assessment")
    .map(l => ({ ...l, isLead: true }));

  // Filter out any leads that are completed or project completed
  const isCompleted = (status: string) => {
    const s = (status || "").toLowerCase().trim();
    return s === "completed" || s === "project completed" || s === "archived";
  };

  // Show all leads marked as Under Assessment in the workspace
  const pendingLeads = assessmentQueueItems.filter(l => !isCompleted(l.leadStatus));

  const [activeTab, setActiveTab] = useState<"pending" | "completed">("pending");

  const [formData, setFormData] = useState<FullAssessmentResult>({
    verification: {
      reputable: { value: false, comment: "" },
      fitsMission: { value: false, comment: "" },
      isNonProfit: { value: false, comment: "" },
      available: { value: false, comment: "" }
    },
    validation: {
      dataQuality: { value: 2, comment: "" },
      problemStatement: { value: 2, comment: "" },
      missionAlignment: { value: 2, comment: "" },
      partnershipReason: { value: 2, comment: "" },
      fundsAvailable: { value: 2, comment: "" }
    },
    validationChecks: {
      diverseStaff: { value: false, comment: "" },
      diverseDemographic: { value: false, comment: "" },
      inclusiveMarketing: { value: false, comment: "" }
    }
  });
  const [extraInfo, setExtraInfo] = useState("");
  const [docUrl, setDocUrl] = useState("");
  const [briefText, setBriefText] = useState("");
  const [summarizing, setSummarizing] = useState(false);
  const [analyzingFile, setAnalyzingFile] = useState(false);
  const [savingBrief, setSavingBrief] = useState(false);
  const [dataAssessment, setDataAssessment] = useState<DataQualityAssessment | null>(null);
  const [assessingData, setAssessingData] = useState(false);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [status, setStatus] = useState<string>("pending");

  useEffect(() => {
    if (selectedNp) {
      const org = leads.find(l => l.id === selectedNp);
      if (org) {
        setDocUrl(org.projectDocUrl || org.website || "");
        setBriefText(org.briefSummary || org.activity || "");
        
        // Load persist data audit from org if available, otherwise fallback to latest assessment
        if (org.dataAssessment) {
          setDataAssessment(org.dataAssessment);
        }

        // Auto-load latest assessment for this partner if in "new" mode but haven't saved anything yet
        if (isNew && !editingId) {
          const latest = assessments
            .filter(a => a.nonProfitId === selectedNp)
            .sort((a, b) => new Date(b.updatedAt || b.createdAt).getTime() - new Date(a.updatedAt || a.createdAt).getTime())[0];
          
          if (latest) {
            setFormData({
              verification: latest.verification || formData.verification,
              validation: latest.validation || formData.validation,
              validationChecks: latest.validationChecks || formData.validationChecks
            });
            setExtraInfo(latest.validationChecks?.extraInfo || "");
            setStatus(latest.status || "pending");
            // Always prefer the org's persisted data audit if available
            setDataAssessment(org.dataAssessment || latest.dataAssessment || null);
          } else {
            resetForm();
            setDataAssessment(org.dataAssessment || null);
          }
        }
      }
    }
  }, [selectedNp, assessments, isNew, editingId]);

  useEffect(() => {
    if (selectedId) {
      if (selectedId.startsWith("new:")) {
        const npId = selectedId.split(":")[1];
        setSelectedNp(npId);
        setIsNew(true);
        setEditingId(null);
        resetForm();
      } else {
        const existing = assessments.find(a => a.id === selectedId);
        if (existing) {
          setEditingId(existing.id);
          setSelectedNp(existing.nonProfitId);
          setFormData({
            verification: existing.verification || formData.verification,
            validation: existing.validation || formData.validation,
            validationChecks: existing.validationChecks || formData.validationChecks
          });
          setExtraInfo(existing.validationChecks?.extraInfo || "");
          setStatus(existing.status || "pending");
          
          // Try to get data assessment from the non-profit first (most recent), fall back to assessment record
          const org = leads.find(l => l.id === existing.nonProfitId);
          setDataAssessment(org?.dataAssessment || existing.dataAssessment || null);
          
          setIsNew(true);
        }
      }
      onClearSelection();
    }
  }, [selectedId, assessments, onClearSelection]);

  const resetForm = () => {
    setFormData({
      verification: {
        reputable: { value: false, comment: "" },
        fitsMission: { value: false, comment: "" },
        isNonProfit: { value: false, comment: "" },
        available: { value: false, comment: "" }
      },
      validation: {
        dataQuality: { value: 2, comment: "" },
        problemStatement: { value: 2, comment: "" },
        missionAlignment: { value: 2, comment: "" },
        partnershipReason: { value: 2, comment: "" },
        fundsAvailable: { value: 2, comment: "" }
      },
      validationChecks: {
        diverseStaff: { value: false, comment: "" },
        diverseDemographic: { value: false, comment: "" },
        inclusiveMarketing: { value: false, comment: "" }
      }
    });
    setExtraInfo("");
    setDataAssessment(null);
    setUploadError(null);
    setStatus("pending");
  };

  const validationValues = Object.values(formData.validation).map(v => (v as any).value as number);
  const totalScore = validationValues.reduce((a, b) => a + b, 0);
  const avg = Number((totalScore / validationValues.length).toFixed(1));

  const handleAutoScore = async (fullTextOverride?: string) => {
    if (!selectedNp) return;
    setScoring(true);
    try {
      const np = leads.find(l => l.id === selectedNp);
      const result = await generateFullAssessment(
        np.name || np.organisation, 
        np.mission || np.activity || "Not specified", 
        np.products_services || np.activity || "Not specified",
        "Expert assessment based on available public data",
        np.briefSummary || np.activity || "",
        { 
          ...aiConfig,
          fullText: fullTextOverride || "" 
        }
      );
      setFormData(result as any);
      notify(fullTextOverride ? "AI Document Scan complete - Rubrics pre-filled" : "AI Assessment complete", "success");
    } catch (err: any) {
      handleAiError(err, "Full Assessment");
    } finally {
      setScoring(false);
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !selectedNp) return;

    setAssessingData(true);
    setUploadError(null);
    setDataAssessment(null);

    const org = leads.find(l => l.id === selectedNp);

    try {
      const reader = new FileReader();
      
      if (file.name.endsWith('.csv')) {
        reader.onload = async (event) => {
          try {
            const text = event.target?.result as string;
            const lines = text.split('\n').slice(0, 50).join('\n');
            const assessment = await assessDataQuality(org.organisation, lines, aiConfig);
            setDataAssessment(assessment);

            await setDoc(doc(db, "leads", selectedNp), {
              dataAssessment: assessment,
              updatedAt: new Date().toISOString()
            }, { merge: true });

            setFormData(prev => ({
              ...prev,
              validation: {
                ...prev.validation,
                dataQuality: { 
                  value: Math.ceil(assessment.score / 2.5), 
                  comment: `AI Data Quality Assessment: ${assessment.verdict}`
                }
              }
            }));
          } catch (err) {
            handleAiError(err, "Data Quality Assessment");
          } finally {
            setAssessingData(false);
          }
        };
        reader.readAsText(file);
      } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        reader.onload = async (event) => {
          try {
            const data = new Uint8Array(event.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            // Take first 50 rows, stringify it
            const snippet = JSON.stringify(json.slice(0, 50));
            const assessment = await assessDataQuality(org.name || org.organisation, snippet, aiConfig);
            setDataAssessment(assessment);

            // Persist data assessment to record
            await setDoc(doc(db, "leads", selectedNp), {
              dataAssessment: assessment,
              updatedAt: new Date().toISOString()
            }, { merge: true });

            setFormData(prev => ({
              ...prev,
              validation: {
                ...prev.validation,
                dataQuality: { 
                  value: Math.ceil(assessment.score / 2.5), 
                  comment: `AI Data Quality Assessment: ${assessment.verdict}`
                }
              }
            }));
          } catch (err) {
            handleAiError(err, "Data Quality Assessment");
          } finally {
            setAssessingData(false);
          }
        };
        reader.readAsArrayBuffer(file);
      } else {
        setUploadError("Unsupported file format. Please upload CSV or Excel.");
        setAssessingData(false);
      }
    } catch (err) {
      console.error(err);
      setUploadError("Failed to parse data file.");
      setAssessingData(false);
    }
  };

  const handleSave = async () => {
    if (!selectedNp) return;
    try {
      const org = leads.find(l => l.id === selectedNp);
      const data = {
        nonProfitId: selectedNp,
        nonProfitName: org?.name || org?.organisation,
        verification: formData.verification,
        validation: formData.validation,
        validationChecks: {
          ...formData.validationChecks,
          extraInfo
        },
        averageScore: avg,
        status: status,
        dataAssessment,
        createdAt: editingId ? assessments.find(a => a.id === editingId)?.createdAt : new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        isValidated: status === "approved"
      };

      if (editingId) {
        await setDoc(doc(db, "assessments", editingId), data);
      } else {
        await addDoc(collection(db, "assessments"), data);
      }

      // Sync status back to CRM Lead if it came from there
      const lead = leads.find(l => l.id === selectedNp);
      if (lead) {
        let newLeadStatus = lead.leadStatus;
        let isVerified = lead.isVerified || false;
        
        if (status === "approved") {
          newLeadStatus = "Approved for Future Project";
          isVerified = true;
        } else if (status === "pending") {
          newLeadStatus = "Under assessment";
        } else if (status === "not suitable") {
          newLeadStatus = "Not suitable";
        } else if (status === "more info") {
          newLeadStatus = "First meeting/contact";
        }

        await setDoc(doc(db, "leads", lead.id), {
          leadStatus: newLeadStatus,
          isVerified: isVerified,
          updatedAt: new Date().toISOString()
        }, { merge: true });
      }
      
      setIsNew(false);
      setEditingId(null);
      resetForm();
      onSetView("briefs", selectedNp);
    } catch (err) {
      handleFirestoreError(err, editingId ? OperationType.UPDATE : OperationType.CREATE, "assessments");
    }
  };

  const handleUpdateBrief = async () => {
    if (!selectedNp) return;
    setSavingBrief(true);
    try {
      let finalDocUrl = docUrl;
      if (finalDocUrl && !finalDocUrl.startsWith('http')) {
        finalDocUrl = 'https://' + finalDocUrl;
      }
      const lead = leads.find(l => l.id === selectedNp);
      await setDoc(doc(db, "leads", selectedNp), { 
        projectDocUrl: finalDocUrl,
        briefSummary: briefText,
        leadStatus: "Assessed",
        assessedBy: userProfile?.email || "Assessor",
        assessedAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      }, { merge: true });
      setDocUrl(finalDocUrl);
      notify("Brief saved and partner moved to 'Assessed'", "success");
    } catch (err) {
      handleFirestoreError(err, OperationType.UPDATE, "leads");
    } finally {
      setSavingBrief(false);
    }
  };

  const handleSummarize = async (textToSummarize?: string) => {
    const rawContent = textToSummarize || briefText;
    if (!selectedNp || !rawContent) return;
    setSummarizing(true);
    try {
      const lead = leads.find(l => l.id === selectedNp);
      const summary = await summarizeBrief(lead.organisation, rawContent, aiConfig);
      setBriefText(summary);
      // Auto-save the summary
      await setDoc(doc(db, "leads", selectedNp), { 
        briefSummary: summary,
        updatedAt: new Date().toISOString() 
      }, { merge: true });
      notify("Project brief summarized", "success");
    } catch (err: any) {
      handleAiError(err, "Brief Summary");
    } finally {
      setSummarizing(false);
    }
  };

  const handleBriefFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !selectedNp) return;
    
    setAnalyzingFile(true);
    try {
      const text = await extractTextFromDoc(file);
      
      // 1. Summarize for the brief field
      await handleSummarize(text);
      
      // 2. Scan for scoring rubrics
      notify("Document uploaded. AI is scanning for verification and validation indicators...", "success");
      await handleAutoScore(text);
      
      notify("Document analyzed, summarized and rubrics pre-filled", "success");
    } catch (err: any) {
      console.error(err);
      notify(err.message || "Failed to analyze document", "error");
    } finally {
      setAnalyzingFile(false);
    }
  };

  const deleteAssessment = async (id: string) => {
    try {
      await deleteDoc(doc(db, "assessments", id));
      if (editingId === id) {
        setIsNew(false);
        setEditingId(null);
        resetForm();
      }
    } catch (err) {
      console.error("Delete assessment error:", err);
      handleFirestoreError(err, OperationType.DELETE, "assessments");
    }
  };

  const handleEdit = (a: any) => {
    setEditingId(a.id);
    setSelectedNp(a.nonProfitId);
    setFormData({
      verification: a.verification || {},
      validation: a.validation || {},
      validationChecks: a.validationChecks || {}
    });
    setExtraInfo(a.validationChecks?.extraInfo || "");
    setIsNew(true);
  };

  return (
    <motion.div 
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      className="max-w-6xl mx-auto"
    >
      <div className="flex justify-between items-center mb-10">
        <div>
          <h2 className="text-xl font-bold text-brand-primary uppercase tracking-tight">Assessment Workspace</h2>
          <p className="text-sm text-brand-muted">Qualifying partners through standard criteria</p>
        </div>
        <div className="flex items-center gap-3">
          {pendingLeads.length > 0 && (
            <div className="animate-bounce flex items-center gap-2 bg-amber-500 text-white px-4 py-2 rounded-full text-[9px] font-black uppercase tracking-widest shadow-lg shadow-amber-500/20">
              <Activity className="w-3 h-3" />
              {pendingLeads.length} NEW PENDING
            </div>
          )}
          {!isNew && (
            <button 
              onClick={() => {
                setEditingId(null);
                setSelectedNp("");
                resetForm();
                setIsNew(true);
              }}
              className="flex items-center gap-2 bg-brand-primary text-white px-6 py-3 rounded-xl font-black text-[10px] uppercase tracking-widest shadow-lg shadow-brand-primary/20 hover:scale-[1.02] transition-all"
            >
              <Plus className="w-4 h-4" /> AD HOC ASSESSMENT
            </button>
          )}
        </div>
      </div>

      {!isNew && (
        <>
          <div className="mb-10 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {pendingLeads.map(lead => {
              const existingAssessment = assessments.find(a => a.nonProfitId === lead.id);
              return (
                <div key={lead.id} className="bg-white border border-brand-border rounded-2xl p-6 card-shadow border-l-4 border-l-amber-500 group hover:border-brand-accent/30 transition-all">
                  <div className="flex justify-between items-start mb-4">
                    <div>
                      <h4 className="text-sm font-black text-brand-primary uppercase truncate w-48">{lead.organisation || lead.name}</h4>
                      {existingAssessment && (
                        <p className="text-[10px] font-bold text-brand-muted uppercase mt-1">
                          Last Updated: {formatDate(existingAssessment.updatedAt || existingAssessment.createdAt, settings?.displayConfig?.dateFormat)}
                        </p>
                      )}
                    </div>
                    <button 
                      onClick={() => {
                        if (existingAssessment) {
                          handleEdit(existingAssessment);
                        } else {
                          setSelectedNp(lead.id);
                          setIsNew(true);
                          setEditingId(null);
                          resetForm();
                        }
                      }}
                      className="p-2 bg-brand-accent text-white rounded-lg shadow-sm hover:scale-105 transition-all"
                      title={existingAssessment ? "Continue Assessment" : "Start Scoring"}
                    >
                      {existingAssessment ? <Target className="w-4 h-4" /> : <BarChart3 className="w-4 h-4" />}
                    </button>
                  </div>
                  <div className="flex items-center justify-between pt-4 border-t border-brand-bg">
                    {existingAssessment ? (
                      <>
                        <span className="text-[9px] font-black text-blue-600 uppercase bg-blue-50 px-2 py-1 rounded">Under Assessment</span>
                        <div className="flex items-center gap-1.5">
                          <span className="text-xs font-black text-brand-primary">{existingAssessment.averageScore}</span>
                          <div className="w-12 h-1 bg-brand-border rounded-full overflow-hidden">
                            <div 
                              className="h-full bg-brand-accent transition-all" 
                              style={{ width: `${(existingAssessment.averageScore / 4) * 100}%` }} 
                            />
                          </div>
                        </div>
                      </>
                    ) : (
                      <span className="text-[9px] font-black text-brand-muted uppercase bg-brand-bg px-2 py-1 rounded">Pending Score</span>
                    )}
                  </div>
                </div>
              );
            })}
          </div>

          <div className="mb-10">
            <h3 className="text-[10px] font-black uppercase tracking-widest text-brand-muted mb-4">Completed Assessments ({assessments.filter(a => {
              const org = leads.find(l => l.id === a.nonProfitId);
              const status = (org?.leadStatus || "").toLowerCase();
              return org && org.leadStatus !== "Under assessment" && !isCompleted(status);
            }).length})</h3>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
              {assessments
                .filter(a => {
                  const org = leads.find(l => l.id === a.nonProfitId);
                  const status = (org?.leadStatus || "").toLowerCase();
                  return org && org.leadStatus !== "Under assessment" && !isCompleted(status);
                })
                .map(a => {
                const org = leads.find(l => l.id === a.nonProfitId);
                return (
                  <div key={a.id} className="bg-white border border-brand-border rounded-2xl p-6 card-shadow border-l-4 border-l-emerald-500 group hover:border-brand-accent/30 transition-all">
                    <div className="flex justify-between items-start mb-4">
                      <div>
                        <h4 className="text-sm font-black text-brand-primary uppercase truncate w-48">{org?.organisation || "Unknown Lead"}</h4>
                        <p className="text-[10px] font-bold text-brand-muted uppercase mt-1">
                          {formatDate(a.updatedAt || a.createdAt, settings?.displayConfig?.dateFormat)}
                        </p>
                      </div>
                      <button 
                        onClick={() => handleEdit(a)}
                        className="p-2 bg-emerald-500 text-white rounded-lg shadow-sm hover:scale-105 transition-all"
                      >
                        <Target className="w-4 h-4" />
                      </button>
                    </div>
                    <div className="flex items-center justify-between pt-4 border-t border-brand-bg">
                      {(() => {
                        let label = a.status === 'approved' ? "APPROVED" : (a.status || 'pending');
                        let color = a.status === 'approved' ? "text-emerald-600" : "text-amber-600";
                        
                        if (org) {
                          if (org.approved) {
                            label = "Approved (Board)";
                            color = "text-emerald-600";
                          } else if (org.needsFurtherAssessment) {
                            label = "Needs Further Work";
                            color = "text-amber-600";
                          } else if ((org.leadStatus || "").toLowerCase().includes("under assessment")) {
                            label = "Under assessment";
                            color = "text-blue-600";
                          } else if ((org.leadStatus || "").toLowerCase() === "assessed") {
                            label = "Assessed";
                            color = "text-emerald-600";
                          }
                        }

                        return (
                          <span className={cn(
                            "text-[9px] font-black uppercase bg-brand-bg px-2 py-1 rounded",
                            color
                          )}>
                            {label}
                          </span>
                        );
                      })()}
                      {isBoardsOrAdmin && (
                        <button 
                          onClick={() => setConfirmingAssessmentId(a.id)}
                          className="p-1.5 text-brand-muted hover:text-red-500 hover:bg-red-50 rounded-lg transition-all"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </>
      )}

      {isNew ? (
        <div className="bg-white border border-brand-border rounded-2xl overflow-hidden card-shadow">
          <div className="bg-brand-primary p-6 flex justify-between items-center text-white">
             <div>
               <h3 className="text-sm font-bold">{editingId ? "Edit Assessment" : "New Assessment Scoring"}</h3>
               <p className="text-[10px] text-white/40 font-bold uppercase tracking-widest mt-1">Status: {avg >= 2.5 ? "Qualifying" : "Review Required"}</p>
             </div>
             <button 
                onClick={() => handleAutoScore()}
                disabled={scoring || !selectedNp}
                className="bg-brand-accent/20 hover:bg-brand-accent/30 text-brand-accent border border-brand-accent/30 px-6 py-2 rounded-lg text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 transition-all"
              >
                {scoring ? <Loader2 className="w-4 h-4 animate-spin" /> : <BarChart3 className="w-4 h-4" />}
                AI Run Full Logic
              </button>
          </div>

          <div className="p-8 space-y-12">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-10 border-b border-brand-border pb-10">
              <div className="space-y-6">
                <SectionHeader label="1. Selected Partner" />
                <select 
                  value={selectedNp}
                  onChange={(e) => setSelectedNp(e.target.value)}
                  className="w-full bg-brand-bg border border-brand-border rounded-xl p-3.5 text-xs font-bold outline-none border-2 border-transparent focus:border-brand-accent transition-all card-shadow"
                >
                  <option value="">Select Lead to Assess</option>
                  {leads && leads.filter(l => (l.leadStatus || "").toLowerCase().includes("under assessment")).length > 0 && (
                    <optgroup label="DCM Leads">
                      {leads.filter(l => (l.leadStatus || "").toLowerCase().includes("under assessment")).map(l => (
                        <option key={l.id} value={l.id}>{l.organisation}</option>
                      ))}
                    </optgroup>
                  )}
                </select>

                <div className={cn(
                  "p-4 border rounded-xl transition-all",
                  analyzingFile ? "bg-amber-50 border-amber-200" : "bg-brand-bg/50 border-brand-border"
                )}>
                  <div className="flex justify-between items-center">
                    <div>
                      <p className="text-[10px] font-bold text-brand-muted uppercase tracking-wider">Project Brief Source</p>
                      <p className="text-[9px] text-brand-muted mt-1 uppercase">AI will scan this doc to pre-fill scoring rubrics</p>
                    </div>
                    <div className="flex items-center gap-3">
                      <label className={cn(
                        "flex items-center gap-2 px-4 py-2 rounded-lg text-[10px] font-black uppercase tracking-widest cursor-pointer transition-all shadow-sm",
                        analyzingFile ? "bg-amber-100 text-amber-700" : "bg-brand-accent text-white hover:scale-105"
                      )}>
                        {analyzingFile ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <UploadCloud className="w-3.5 h-3.5" />}
                        {analyzingFile ? "Analyzing Brief..." : "Upload & AI Scan Brief"}
                        <input type="file" className="hidden" accept=".pdf,.docx,text/plain" onChange={handleBriefFileUpload} disabled={analyzingFile || !selectedNp} />
                      </label>
                    </div>
                  </div>
                </div>
              </div>

              <div className="space-y-6">
                <SectionHeader label="2. Strategic Summary distilling Objectives" />
                <div className="relative group">
                  <textarea 
                    value={briefText}
                    onChange={(e) => setBriefText(e.target.value)}
                    placeholder="Paste the charity's project brief or sample data summary here..."
                    className="w-full h-40 bg-brand-bg border border-brand-border rounded-xl p-4 text-xs font-medium outline-none focus:ring-2 focus:ring-brand-accent/20 resize-none transition-all shadow-sm"
                  />
                  <div className="absolute bottom-3 right-3 flex gap-2 opacity-0 group-hover:opacity-100 transition-all">
                    <button 
                      onClick={() => handleSummarize()}
                      disabled={summarizing || !briefText}
                      className="p-2 bg-brand-accent text-white rounded-lg opacity-80 hover:opacity-100 disabled:opacity-30 transition-all shadow-sm"
                      title="AI Summarize"
                    >
                      {summarizing ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <BarChart3 className="w-3.5 h-3.5" />}
                    </button>
                    <button 
                      onClick={handleUpdateBrief}
                      disabled={savingBrief || !selectedNp}
                      className="p-2 bg-brand-primary text-white rounded-lg opacity-80 hover:opacity-100 disabled:opacity-30 transition-all shadow-sm"
                      title="Save Brief"
                    >
                      {savingBrief ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Save className="w-3.5 h-3.5" />}
                    </button>
                  </div>
                </div>
              </div>
            </div>

            <div className="border-b border-brand-border pb-10">
                <SectionHeader label="3. RAW Sample Data Set" />
                <div className={cn(
                  "border-2 border-dashed border-brand-border rounded-2xl p-8 transition-all hover:border-brand-accent/40 bg-brand-bg/20 flex flex-col items-center justify-center h-48",
                  !selectedNp && "opacity-50 pointer-events-none"
                )}>
                  <label className="flex flex-col items-center gap-3 cursor-pointer w-full h-full justify-center">
                    <div className="w-12 h-12 bg-white rounded-2xl shadow-md flex items-center justify-center border border-brand-border group-hover:scale-110 transition-all">
                      {assessingData ? <Loader2 className="w-5 h-5 text-brand-accent animate-spin" /> : <UploadCloud className="w-6 h-6 text-brand-accent" />}
                    </div>
                    <div className="text-center">
                      <p className="text-xs font-bold text-brand-primary mb-1 uppercase tracking-tight">Upload Raw Data Sample</p>
                      <p className="text-[10px] font-medium text-brand-muted uppercase tracking-[0.2em]">
                        {assessingData ? "AI Audit in progress..." : "CSV, XLSX supported"}
                      </p>
                    </div>
                    <input type="file" className="hidden" accept=".csv, .xlsx, .xls" onChange={handleFileUpload} />
                  </label>
                </div>
                {uploadError && <p className="text-[10px] text-red-500 font-bold mt-2 uppercase tracking-tight">{uploadError}</p>}
            </div>

            {dataAssessment && (
              <motion.div 
                initial={{ opacity: 0, scale: 0.98 }}
                animate={{ opacity: 1, scale: 1 }}
                className="bg-brand-primary p-8 rounded-2xl text-white relative overflow-hidden"
              >
                <div className="absolute top-0 right-0 p-8 opacity-[0.05] pointer-events-none">
                  <BarChart3 className="w-48 h-48" />
                </div>
                
                <div className="relative">
                  <div className="flex justify-between items-start mb-8">
                    <div>
                      <h4 className="text-2xl font-black mb-1">AI Data Quality Audit</h4>
                      <p className="text-[10px] uppercase font-bold text-white/40 tracking-[0.2em]">Partner Provided Sample Assessment</p>
                    </div>
                    <div className="text-right">
                      <p className="text-[10px] uppercase font-bold text-white/40 tracking-widest mb-1">Visualisation Score</p>
                      <p className="text-5xl font-black text-brand-accent tabular-nums">{dataAssessment.score}<span className="text-xl text-white/20">/10</span></p>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-10">
                    <div className="space-y-6">
                      <div className="p-4 bg-white/5 border border-white/10 rounded-xl italic text-xs leading-relaxed text-white/70">
                        "{dataAssessment.verdict}"
                      </div>
                      
                      <div>
                        <SectionHeader label="Sample Strengths" className="text-white/40" />
                        <ul className="space-y-2">
                          {dataAssessment.strengths.map((s, i) => (
                            <li key={i} className="flex gap-2 text-xs font-medium text-white/80">
                              <span className="text-emerald-400">✓</span> {s}
                            </li>
                          ))}
                        </ul>
                      </div>
                    </div>

                    <div className="space-y-6">
                      <div>
                        <SectionHeader label="Data Weaknesses" className="text-white/40" />
                        <ul className="space-y-2">
                          {dataAssessment.weaknesses.map((w, i) => (
                            <li key={i} className="flex gap-2 text-xs font-medium text-white/80">
                              <span className="text-amber-400">!</span> {w}
                            </li>
                          ))}
                        </ul>
                      </div>

                      <div>
                        <SectionHeader label="Expert Recommendations" className="text-white/40" />
                        <ul className="space-y-2">
                          {dataAssessment.recommendations.map((r, i) => (
                            <li key={i} className="flex gap-2 text-xs font-medium text-white/80">
                              <span className="text-brand-accent">•</span> {r}
                            </li>
                          ))}
                        </ul>
                      </div>
                    </div>
                  </div>
                </div>
              </motion.div>
            )}

            {/* Intelligence Rubrics */}
            {(settings?.scoring?.verification || []).length > 0 && (
              <div className="space-y-4">
                <RubricHeader label="1. Verification" description="Foundational alignment and status" />
                <div className="border border-brand-border rounded-xl divide-y divide-brand-border overflow-hidden">
                  {settings.scoring.verification.map((q: any) => (
                    <RubricRow 
                      key={q.id}
                      label={q.label}
                      type={q.type}
                      value={formData.verification[q.id]?.value}
                      comment={formData.verification[q.id]?.comment}
                      onChange={(v) => setFormData(prev => ({
                        ...prev,
                        verification: { ...prev.verification, [q.id]: { ...prev.verification[q.id], value: v } }
                      }))}
                      onCommentChange={(c) => setFormData(prev => ({
                        ...prev,
                        verification: { ...prev.verification, [q.id]: { ...prev.verification[q.id], comment: c } }
                      }))}
                      hint={q.hint}
                    />
                  ))}
                </div>
              </div>
            )}

            {(settings?.scoring?.validation || []).length > 0 && (
              <div className="space-y-4">
                <RubricHeader label="2. Validation" description="Numerical scoring against key impact metrics" />
                <div className="border border-brand-border rounded-xl divide-y divide-brand-border overflow-hidden">
                  {settings.scoring.validation.map((q: any) => (
                    <RubricRow 
                      key={q.id}
                      label={q.label}
                      type={q.type}
                      value={formData.validation[q.id]?.value ?? (q.type === 'score' ? 2 : false)}
                      comment={formData.validation[q.id]?.comment || ""}
                      onChange={(v) => setFormData(prev => ({
                        ...prev,
                        validation: { ...prev.validation, [q.id]: { ...prev.validation[q.id], value: v } }
                      }))}
                      onCommentChange={(c) => setFormData(prev => ({
                        ...prev,
                        validation: { ...prev.validation, [q.id]: { ...prev.validation[q.id], comment: c } }
                      }))}
                    />
                  ))}
                  
                  <div className="bg-brand-bg/30 px-8 py-6 flex justify-between items-center">
                    <div>
                      <p className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Aggregate Validation Result</p>
                      <p className="text-xs font-medium text-brand-muted mt-1 italic">If score is less than 2, org might not qualify</p>
                    </div>
                    <div className="flex items-center gap-8">
                      <div className="text-right">
                        <p className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Total</p>
                        <p className="text-xl font-black text-brand-primary leading-none">{totalScore}</p>
                      </div>
                      <div className="text-right">
                        <p className="text-[10px] font-bold text-brand-accent uppercase tracking-widest">Average</p>
                        <p className="text-4xl font-black text-brand-accent leading-none">{avg}</p>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {(settings?.scoring?.validationChecks || []).length > 0 && (
              <div className="space-y-4">
                <RubricHeader label="3. Validation Checks" description="Diverse demographic and inclusive marketing signals" />
                <div className="border border-brand-border rounded-xl divide-y divide-brand-border overflow-hidden">
                  {settings.scoring.validationChecks.map((q: any) => (
                    <RubricRow 
                      key={q.id}
                      label={q.label}
                      type={q.type}
                      value={formData.validationChecks[q.id]?.value}
                      comment={formData.validationChecks[q.id]?.comment}
                      onChange={(v) => setFormData(prev => ({
                        ...prev,
                        validationChecks: { ...prev.validationChecks, [q.id]: { ...prev.validationChecks[q.id], value: v } }
                      }))}
                      onCommentChange={(c) => setFormData(prev => ({
                        ...prev,
                        validationChecks: { ...prev.validationChecks, [q.id]: { ...prev.validationChecks[q.id], comment: c } }
                      }))}
                    />
                  ))}
                  
                  <div className="p-8 bg-brand-bg/10">
                    <h5 className="text-[10px] font-bold text-brand-muted uppercase tracking-widest mb-3">Ask for extra info if required</h5>
                    <textarea 
                      value={extraInfo}
                      onChange={(e) => setExtraInfo(e.target.value)}
                      placeholder="Capture additional context here..."
                      className="w-full h-24 bg-white border border-brand-border rounded-lg p-4 text-sm font-medium outline-none focus:ring-2 focus:ring-brand-accent/20 transition-all resize-none"
                    />
                  </div>
                </div>
              </div>
            )}

            {/* 4. Outcome Section */}
            <div className="space-y-6 pt-10 border-t border-brand-border">
              <RubricHeader label="4. Final Outcome" description="Set project status and review consolidated insights" />
              
              <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
                <div className="md:col-span-2 bg-brand-bg/20 border border-brand-border rounded-2xl p-8">
                  <h5 className="text-[10px] font-black uppercase tracking-wider text-brand-muted mb-6">Consolidated Intelligence Summary</h5>
                  <div className="space-y-6">
                    <div className="flex gap-4">
                      <div className="w-1 bg-brand-accent rounded-full" />
                      <div>
                        <p className="text-[10px] font-bold text-brand-primary uppercase">Calculated Quality Index</p>
                        <p className="text-[13px] font-medium text-brand-text mt-1">Based on scoring, this partner has an average validation of <span className="font-bold text-brand-accent">{avg} / 4</span>.</p>
                      </div>
                    </div>
                    <div className="flex gap-4">
                      <div className={cn("w-1 rounded-full", dataAssessment ? "bg-emerald-500" : "bg-brand-border")} />
                      <div>
                        <p className="text-[10px] font-bold text-brand-primary uppercase">Technical Readiness</p>
                        <p className="text-[13px] font-medium text-brand-text mt-1">
                          {dataAssessment 
                            ? `Data sample has a visualisation score of ${dataAssessment.score}/10: ${dataAssessment.verdict}`
                            : "No sample data has been assessed via the technical upload yet."}
                        </p>
                      </div>
                    </div>
                    <div className="flex gap-4">
                      <div className={cn("w-1 rounded-full", leads.find(l => l.id === selectedNp)?.briefSummary ? "bg-brand-accent" : "bg-brand-border")} />
                      <div>
                        <p className="text-[10px] font-bold text-brand-primary uppercase">Mission Alignment</p>
                        <p className="text-[13px] font-medium text-brand-text mt-1">
                          {leads.find(l => l.id === selectedNp)?.briefSummary 
                            ? "Strategy brief successfully distilled via Agent Insight."
                            : "No strategic brief summary currently extracted for this partner."}
                        </p>
                      </div>
                    </div>
                  </div>
                </div>

                {isBoardsOrAdmin ? (
                  <div className="bg-white border-2 border-brand-primary rounded-2xl p-8 card-shadow shadow-brand-primary/10">
                    <h5 className="text-[10px] font-black uppercase tracking-wider text-brand-primary mb-6">Decision Outcome</h5>
                    <div className="space-y-6">
                      <div className="space-y-3">
                        <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Select Project Status</label>
                        <select 
                          value={status}
                          onChange={(e) => setStatus(e.target.value)}
                          className={cn(
                            "w-full p-4 rounded-xl font-bold text-xs uppercase tracking-widest border-2 outline-none transition-all",
                            status === "approved" ? "border-emerald-500 bg-emerald-50 text-emerald-700" :
                            status === "pending" ? "border-amber-500 bg-amber-50 text-amber-700" :
                            status === "not suitable" ? "border-rose-500 bg-rose-50 text-rose-700" :
                            status === "more info" ? "border-blue-500 bg-blue-50 text-blue-700" : "border-brand-border bg-brand-bg text-brand-primary"
                          )}
                        >
                          {settings?.leadsConfig?.statuses?.map((s: string) => (
                            <option key={s} value={s.toLowerCase() === "under assessment" ? "pending" : s.toLowerCase() === "approved for future project" ? "approved" : s}>{s}</option>
                          )) || (
                            <>
                              <option value="pending">Under Assessment</option>
                              <option value="approved">Approved for Project</option>
                              <option value="more info">Needs More Information</option>
                              <option value="not suitable">Not Suitable</option>
                            </>
                          )}
                        </select>
                      </div>
                      
                      <div className="p-4 bg-brand-bg rounded-xl border border-brand-border">
                         <p className="text-[10px] font-medium text-brand-muted italic leading-relaxed">
                           Setting an outcome will update the Discovery pipeline and potentially unlock the Recommendation Pack generation for the partner.
                         </p>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="bg-brand-bg p-8 rounded-2xl border border-brand-border">
                    <h5 className="text-[10px] font-black uppercase tracking-wider text-brand-muted mb-4">Outcome</h5>
                    <div className="p-4 bg-white/50 rounded-xl border border-brand-border/50">
                       <p className="text-[10px] font-medium text-brand-muted italic leading-relaxed">
                         Assessments are reviewed by the DCM Board. Finalize this record to move to the Assessment Pack generation phase.
                       </p>
                    </div>
                  </div>
                )}
              </div>
            </div>

            <div className="flex gap-4 pt-10 border-t border-brand-border">
              <button 
                onClick={handleSave}
                disabled={userProfile?.role?.toLowerCase() === "dcm boards"}
                className="px-10 py-3.5 bg-brand-accent text-white font-bold text-xs uppercase tracking-widest rounded-lg hover:bg-brand-accent/90 transition-all shadow-lg shadow-brand-accent/20 disabled:opacity-30 disabled:cursor-not-allowed"
              >
                FINALIZE ASSESSMENT
              </button>
              <button 
                onClick={() => setIsNew(false)}
                className="px-10 py-3.5 bg-brand-bg text-brand-muted font-bold text-xs uppercase tracking-widest rounded-lg border border-brand-border hover:bg-white transition-all"
              >
                DISCARD
              </button>
            </div>
          </div>
        </div>
      ) : (

        <div className="bg-white border border-brand-border rounded-xl card-shadow overflow-hidden">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-brand-bg text-brand-muted font-bold text-[10px] uppercase tracking-widest border-b border-brand-border">
                <th className="px-8 py-5">Organization</th>
                <th className="px-8 py-5">Date Logged</th>
                <th className="px-8 py-5 text-center">Assessment Score</th>
                <th className="px-8 py-5">Outcome</th>
                <th className="px-8 py-5"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-brand-border">
              {assessments
                .filter(a => {
                  const org = leads.find(l => l.id === a.nonProfitId);
                  return org && !isCompleted(org.leadStatus);
                })
                .map(a => (
                <tr key={a.id} className="hover:bg-brand-bg/40 transition-colors group">
                  <td className="px-8 py-6">
                    <div className="flex items-center gap-3">
                      <div className="w-8 h-8 rounded-lg bg-blue-50 text-brand-accent flex items-center justify-center font-bold text-[10px]">
                        {(String(a.nonProfitName || "NP")).substring(0, 2).toUpperCase()}
                      </div>
                      <span className="font-bold text-brand-primary">{a.nonProfitName || "Assessment"}</span>
                    </div>
                  </td>
                  <td className="px-8 py-6 text-[11px] font-medium text-brand-muted">
                    {formatDate(a.createdAt, settings?.displayConfig?.dateFormat)}
                  </td>
                  <td className="px-8 py-6 text-center">
                    <div className="flex flex-col items-center gap-1.5">
                      <span className="text-sm font-bold text-brand-primary leading-none">{a.averageScore}</span>
                      <div className="w-16 h-1 bg-brand-border rounded-full overflow-hidden">
                        <div 
                          className={cn("h-full transition-all", a.averageScore >= 3 ? "bg-brand-accent" : "bg-amber-400")} 
                          style={{ width: `${(a.averageScore / 4) * 100}%` }} 
                        />
                      </div>
                    </div>
                  </td>
                  <td className="px-8 py-6">
                    {(() => {
                      const org = leads.find(l => l.id === a.nonProfitId);
                      let label = a.status === "approved" ? "APPROVED READY" : 
                                  a.status === "pending" ? "IN REVIEW" :
                                  a.status === "not suitable" ? "NOT SUITABLE" : "MORE INFO";
                      let style = a.status === "approved" ? "text-emerald-700 bg-emerald-50 border-emerald-100" : 
                                  a.status === "pending" ? "text-amber-700 bg-amber-50 border-amber-100" :
                                  a.status === "not suitable" ? "text-rose-700 bg-rose-50 border-rose-100" : "text-blue-700 bg-blue-50 border-blue-100";

                      if (org) {
                        if (org.approved) {
                          label = "Approved (Board)";
                          style = "text-emerald-700 bg-emerald-50 border-emerald-100";
                        } else if (org.needsFurtherAssessment) {
                          label = "Needs Further Work";
                          style = "text-amber-700 bg-amber-50 border-amber-100";
                        } else if ((org.leadStatus || "").toLowerCase().includes("under assessment")) {
                          label = "Under assessment";
                          style = "text-blue-700 bg-blue-50 border-blue-100";
                        } else if ((org.leadStatus || "").toLowerCase() === "assessed") {
                          label = "Assessed";
                          style = "text-emerald-700 bg-emerald-50 border-emerald-100";
                        }
                      }

                      return (
                        <span className={cn(
                          "text-[9px] font-bold uppercase tracking-widest px-2.5 py-1 rounded-md border",
                          style
                        )}>
                          {label}
                        </span>
                      );
                    })()}
                  </td>
                  <td className="px-8 py-6 text-right">
                    <div className="flex items-center justify-end gap-2">
                      <button 
                        onClick={() => handleEdit(a)}
                        className="p-2 text-brand-muted hover:text-brand-accent transition-all flex items-center gap-1 group/btn"
                      >
                        <span className="text-[9px] font-bold uppercase tracking-widest opacity-0 group-hover/btn:opacity-100 transition-all">Edit</span>
                        <ChevronRight className="w-4 h-4" />
                      </button>
                      {isBoardsOrAdmin && (
                        confirmingAssessmentId === a.id ? (
                          <div className="flex items-center gap-1 bg-red-50 p-1 rounded-lg border border-red-100">
                            <button 
                              onClick={() => {
                                deleteAssessment(a.id);
                                setConfirmingAssessmentId(null);
                              }}
                              className="text-[10px] font-bold text-red-600 px-2 py-1 hover:bg-red-100 rounded transition-all"
                            >
                              Confirm?
                            </button>
                            <button 
                              onClick={() => setConfirmingAssessmentId(null)}
                              className="p-1 text-slate-400 hover:text-slate-600 rounded"
                            >
                              <X className="w-3 h-3" />
                            </button>
                          </div>
                        ) : (
                          <button 
                            onClick={() => setConfirmingAssessmentId(a.id)}
                            className="p-2 text-brand-muted hover:text-red-500 transition-all"
                            title="Delete Assessment"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        )
                      )}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {assessments.length === 0 && (
            <div className="p-12 text-center">
              <p className="text-sm text-brand-muted">No assessments recorded in this portfolio yet.</p>
            </div>
          )}
        </div>
      )}
    </motion.div>
  );
}

function BriefsView({ assessments, leads, settings, userProfile, aiConfig, onShowAiConfig, handleAiError, notify, initialLeadId, onClearSelection, logLeadEvent }: { 
  assessments: any[], 
  leads: any[], 
  settings: any, 
  userProfile: any, 
  aiConfig: any,
  onShowAiConfig: () => void,
  handleAiError: (err: any, source: string) => void,
  notify: (msg: string, type: 'success' | 'error') => void,
  initialLeadId?: string | null,
  onClearSelection?: () => void,
  logLeadEvent: (leadId: string, type: string, description: string, metadata?: any) => Promise<void>
}) {
  const [selectedNp, setSelectedNp] = useState(initialLeadId || "");
  const [showRecommendation, setShowRecommendation] = useState(false);
  const [isConsolidating, setIsConsolidating] = useState(false);
  const [consolidatedData, setConsolidatedData] = useState<any>(null);
  const [recommendationSaved, setRecommendationSaved] = useState(false);
  
  const selectedOrg = leads.find(l => l.id === selectedNp);
  const orgAssessment = assessments.find(a => a.nonProfitId === selectedNp);

  useEffect(() => {
    if ((selectedOrg?.leadStatus || "").toLowerCase() === "assessed") {
      setRecommendationSaved(true);
    } else {
      setRecommendationSaved(false);
    }
    setIsConsolidating(false);
  }, [selectedNp, selectedOrg?.leadStatus]);

  useEffect(() => {
    if (initialLeadId) {
      setSelectedNp(initialLeadId);
      onClearSelection?.();
    }
  }, [initialLeadId]);

  const [lastRefreshedLeadId, setLastRefreshedLeadId] = useState<string | null>(null);

  useEffect(() => {
    if (selectedOrg) {
      // If the lead ID changed, force a full refresh of consolidated data
      if (selectedOrg.id !== lastRefreshedLeadId) {
        setConsolidatedData({
          briefSummary: selectedOrg.briefSummary || "",
          revenue: selectedOrg.revenue || "",
          ein: selectedOrg.ein || "",
          charity_navigator_rating: selectedOrg.charity_navigator_rating || "",
          propublica_grants: selectedOrg.propublica_grants || "",
          linkedin_overview: selectedOrg.linkedin_overview || selectedOrg.linkedin_activity || "",
          staff_linkedin_summary: selectedOrg.staff_linkedin_summary || selectedOrg.staff_members || "",
          dcmComment: selectedOrg.dcmComment || "",
          dcmName: selectedOrg.dcmName || ""
        });
        setLastRefreshedLeadId(selectedOrg.id);
      } else if (!isConsolidating && !consolidatedData) {
        // Initial load for same lead
        setConsolidatedData({
          briefSummary: selectedOrg.briefSummary || "",
          revenue: selectedOrg.revenue || "",
          ein: selectedOrg.ein || "",
          charity_navigator_rating: selectedOrg.charity_navigator_rating || "",
          propublica_grants: selectedOrg.propublica_grants || "",
          linkedin_overview: selectedOrg.linkedin_overview || selectedOrg.linkedin_activity || "",
          staff_linkedin_summary: selectedOrg.staff_linkedin_summary || selectedOrg.staff_members || "",
          dcmComment: selectedOrg.dcmComment || "",
          dcmName: selectedOrg.dcmName || ""
        });
      }
    } else {
      setConsolidatedData(null);
      setLastRefreshedLeadId(null);
    }
  }, [selectedOrg, isConsolidating, lastRefreshedLeadId, consolidatedData]);

  const updatePackField = (field: string, val: string) => {
    setConsolidatedData((prev: any) => ({ ...prev, [field]: val }));
  };

  const handleUpdateBrief = async () => {
    if (!selectedNp) return;
    setIsConsolidating(true); // Re-use indicator for save
    try {
      const org = leads.find(l => l.id === selectedNp);
      
      // Clean updates: only include significant changes to avoid wiping existing research
      const updates: any = { 
        updatedAt: new Date().toISOString(),
        leadStatus: "Assessed",
        assessedBy: userProfile?.email || "Assessor",
        assessedAt: new Date().toISOString(),
        needsFurtherAssessment: false // Clear the flag when assessment is updated
      };

      // Map the pack fields
      if (consolidatedData) {
        Object.keys(consolidatedData).forEach(key => {
          const val = consolidatedData[key];
          const lowerVal = String(val || "").toLowerCase();
          const isNotAvailable = !val || 
                                lowerVal === "" || 
                                lowerVal.includes("no public record found") || 
                                lowerVal.includes("not found");

          // Only update if we have actual content
          if (!isNotAvailable) {
            updates[key] = val;
          }
        });
      }
      
      console.log("Updating lead status to Assessed for:", selectedNp);
      await setDoc(doc(db, "leads", selectedNp), updates, { merge: true });

      // Log to timeline if there's a DCM comment
      if (consolidatedData?.dcmComment) {
        await logLeadEvent(
          selectedNp,
          "Note",
          `DCM Recommendation Added by ${consolidatedData.dcmName || userProfile?.email || 'Assessor'}: ${consolidatedData.dcmComment}`,
          { source: "Assessment Pack" }
        );
      }

      setRecommendationSaved(true);
      notify("Recommendation saved and lead moved to 'Assessed'", "success");
    } catch (err) {
      handleFirestoreError(err, OperationType.UPDATE, "leads");
    } finally {
      setIsConsolidating(false);
    }
  };

  return (
    <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="max-w-6xl mx-auto">
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-10">
        <div className="space-y-6">
          <div className="bg-white border border-brand-border rounded-xl p-8 card-shadow text-brand-primary">
            <SectionHeader label="1. Select CRM Lead" />
            <select 
              value={selectedNp}
              onChange={(e) => setSelectedNp(e.target.value)}
              className="w-full bg-brand-bg border border-brand-border rounded-lg p-3 text-sm font-bold outline-none focus:ring-2 focus:ring-brand-accent/20"
            >
              <option value="">Select Assessment lead...</option>
              {leads && leads.filter(l => (l.leadStatus || "").toLowerCase().includes("under assessment")).length > 0 && (
                <optgroup label="Active Pipeline (Under assessment)">
                  {leads.filter(l => (l.leadStatus || "").toLowerCase().includes("under assessment")).map(l => (
                    <option key={l.id} value={l.id}>{l.organisation} {l.isVerified ? '✓' : ''}</option>
                  ))}
                </optgroup>
              )}
              {leads && leads.filter(l => (l.leadStatus || "").toLowerCase() === "assessed").length > 0 && (
                <optgroup label="Recently Assessed (Needs Review)">
                  {leads.filter(l => (l.leadStatus || "").toLowerCase() === "assessed").map(l => (
                    <option key={l.id} value={l.id}>{l.organisation} {l.isVerified ? '✓' : ''}</option>
                  ))}
                </optgroup>
              )}
            </select>

            <div className="mt-10 pt-10 border-t border-brand-border space-y-4">
               {!recommendationSaved ? (
                 <button 
                   onClick={() => setIsConsolidating(true)}
                   disabled={!selectedNp || !orgAssessment}
                   className="w-full py-4 bg-brand-primary text-white rounded-xl font-black text-[11px] uppercase tracking-widest hover:bg-brand-primary/90 disabled:opacity-50 shadow-lg shadow-brand-primary/20 flex items-center justify-center gap-2"
                 >
                    <ShieldCheck className="w-4 h-4 text-brand-accent" /> PREPARE RECOMMENDATIONS
                 </button>
               ) : (
                 <button 
                   onClick={() => setShowRecommendation(true)}
                   className="w-full py-4 bg-brand-accent text-white rounded-xl font-black text-[11px] uppercase tracking-widest hover:bg-brand-accent/90 shadow-lg shadow-brand-accent/20 flex items-center justify-center gap-2"
                 >
                    <ArrowRightCircle className="w-4 h-4" /> GENERATE ASSESSMENT PACK
                 </button>
               )}
               
              {!orgAssessment && selectedNp && (
                <p className="text-[10px] text-rose-600 font-bold mt-4 text-center">Assessment required before pack generation</p>
              )}

              {isConsolidating && (
                <button 
                  onClick={() => setIsConsolidating(false)}
                  className="w-full py-2 text-[10px] font-bold text-brand-muted uppercase tracking-widest hover:text-brand-primary"
                >
                  Cancel Preparation
                </button>
              )}
            </div>
          </div>
          
          <div className="bg-brand-primary p-8 rounded-xl text-white">
            <h4 className="text-xl font-black mb-4 flex items-center gap-2">
              <FileText className="w-5 h-5 text-brand-accent" />
              Agent Insight
            </h4>
            <div className="bg-white/5 p-4 rounded-xl border border-white/10 markdown-content text-white/90">
                <Markdown>
                  {String(selectedOrg?.briefSummary ? selectedOrg.briefSummary : '"Consolidate the partner records to view the final intelligence summary here."')}
                </Markdown>
            </div>
          </div>
        </div>

        <div className="lg:col-span-2">
          {isConsolidating && consolidatedData ? (
            <motion.div 
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              className="bg-white border border-brand-border rounded-2xl p-10 card-shadow space-y-10"
            >
              <div className="pb-6 border-b border-brand-border flex justify-between items-center">
                <div>
                  <h3 className="text-2xl font-black text-brand-primary uppercase tracking-tighter">Assessment pack</h3>
                  <p className="text-sm text-brand-muted font-medium">Refine findings from AI Research, Scoring, and Assessment before final generation.</p>
                </div>
                <button 
                  onClick={() => {
                    if (!selectedOrg) return;
                    setConsolidatedData({
                      briefSummary: selectedOrg.briefSummary || "",
                      revenue: selectedOrg.revenue || "",
                      ein: selectedOrg.ein || "",
                      charity_navigator_rating: selectedOrg.charity_navigator_rating || "",
                      propublica_grants: selectedOrg.propublica_grants || "",
                      linkedin_overview: selectedOrg.linkedin_overview || selectedOrg.linkedin_activity || "",
                      staff_linkedin_summary: selectedOrg.staff_linkedin_summary || selectedOrg.staff_members || "",
                      dcmComment: selectedOrg.dcmComment || "",
                      dcmName: selectedOrg.dcmName || ""
                    });
                    notify("Data flushed and synchronized with Lead record.", "success");
                  }}
                  className="flex items-center gap-2 px-4 py-2 bg-rose-50 text-rose-600 border border-rose-100 rounded-lg text-[10px] font-black uppercase tracking-widest hover:bg-rose-100 transition-all"
                  title="Overwrite local edits with latest data from Research/Leads record"
                >
                  <RotateCw className="w-3 h-3" /> Flush & Refresh
                </button>
              </div>

              <div className="space-y-8">
                <div>
                  <label className="text-[10px] font-black uppercase tracking-widest text-brand-primary mb-3 block">Strategic Discovery Summary</label>
                  <textarea 
                    value={consolidatedData.briefSummary}
                    onChange={(e) => updatePackField('briefSummary', e.target.value)}
                    className="w-full h-32 bg-brand-bg border border-brand-border rounded-xl p-4 text-xs font-medium outline-none focus:ring-2 focus:ring-brand-accent/20 resize-none"
                  />
                </div>

                <div className="grid grid-cols-2 gap-6">
                  <div>
                    <label className="text-[10px] font-black uppercase tracking-widest text-brand-primary mb-3 block">Revenue / Financial Tier</label>
                    <input 
                      value={consolidatedData.revenue}
                      onChange={(e) => updatePackField('revenue', e.target.value)}
                      className="w-full bg-brand-bg border border-brand-border rounded-lg p-3 text-xs font-bold"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] font-black uppercase tracking-widest text-brand-primary mb-3 block">EIN Status</label>
                    <input 
                      value={consolidatedData.ein}
                      onChange={(e) => updatePackField('ein', e.target.value)}
                      className="w-full bg-brand-bg border border-brand-border rounded-lg p-3 text-xs font-bold"
                    />
                  </div>
                </div>

                <div>
                  <label className="text-[10px] font-black uppercase tracking-widest text-brand-primary mb-3 block">Charity Navigator Findings</label>
                  <textarea 
                    value={consolidatedData.charity_navigator_rating}
                    onChange={(e) => updatePackField('charity_navigator_rating', e.target.value)}
                    className="w-full h-24 bg-brand-bg border border-brand-border rounded-xl p-4 text-xs font-medium outline-none focus:ring-2 focus:ring-brand-accent/20 resize-none"
                  />
                </div>

                <div>
                  <label className="text-[10px] font-black uppercase tracking-widest text-brand-primary mb-3 block">ProPublica Grant Intelligence</label>
                  <textarea 
                    value={consolidatedData.propublica_grants}
                    onChange={(e) => updatePackField('propublica_grants', e.target.value)}
                    className="w-full h-24 bg-brand-bg border border-brand-border rounded-xl p-4 text-xs font-medium outline-none focus:ring-2 focus:ring-brand-accent/20 resize-none"
                  />
                </div>

                <div className="grid grid-cols-2 gap-6">
                  <div>
                    <label className="text-[10px] font-black uppercase tracking-widest text-brand-primary mb-3 block">LinkedIn Audit Overview</label>
                    <textarea 
                      value={consolidatedData.linkedin_overview}
                      onChange={(e) => updatePackField('linkedin_overview', e.target.value)}
                      className="w-full h-32 bg-brand-bg border border-brand-border rounded-xl p-4 text-xs font-medium outline-none focus:ring-2 focus:ring-brand-accent/20 resize-none"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] font-black uppercase tracking-widest text-brand-primary mb-3 block">Leadership Blueprint</label>
                    <textarea 
                      value={consolidatedData.staff_linkedin_summary}
                      onChange={(e) => updatePackField('staff_linkedin_summary', e.target.value)}
                      className="w-full h-32 bg-brand-bg border border-brand-border rounded-xl p-4 text-xs font-medium outline-none focus:ring-2 focus:ring-brand-accent/20 resize-none"
                    />
                  </div>
                </div>

                <div className="pt-6 border-t border-brand-border space-y-6">
                  <SectionHeader label="Executive Recommendation (DCM)" />
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div>
                      <label className="text-[10px] font-black uppercase tracking-widest text-brand-primary mb-3 block">DCM Comment / Strategic Recommendation</label>
                      <textarea 
                        value={consolidatedData.dcmComment}
                        onChange={(e) => updatePackField('dcmComment', e.target.value)}
                        placeholder="Provide executive guidance for DCM stakeholders..."
                        className="w-full h-32 bg-brand-bg border border-brand-border rounded-xl p-4 text-xs font-medium outline-none focus:ring-2 focus:ring-brand-accent/20 resize-none"
                      />
                    </div>
                    <div>
                      <label className="text-[10px] font-black uppercase tracking-widest text-brand-primary mb-3 block">DCM Lead / Consultant Name</label>
                      <input 
                        value={consolidatedData.dcmName}
                        onChange={(e) => updatePackField('dcmName', e.target.value)}
                        placeholder="Your name..."
                        className="w-full bg-brand-bg border border-brand-border rounded-lg p-3 text-xs font-bold"
                      />
                      <button 
                        onClick={handleUpdateBrief}
                        className="mt-6 w-full py-3 bg-brand-accent/10 text-brand-accent border border-brand-accent/20 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-brand-accent/20 flex items-center justify-center gap-2 transition-all"
                      >
                        <Save className="w-4 h-4" /> Save Recommendation Changes
                      </button>
                      <p className="mt-4 text-[9px] font-bold text-brand-muted uppercase tracking-widest text-center leading-relaxed">
                        Note: After saving, please generate the <span className="text-brand-accent">Assessment Pack</span> below and save the final PDF to the Lead records.
                      </p>
                    </div>
                  </div>
                </div>
              </div>
            </motion.div>
          ) : (
            <div className="bg-white border border-brand-border rounded-2xl p-16 card-shadow relative overflow-hidden min-h-[900px]">
              <div className="absolute top-0 right-0 p-12 opacity-[0.03] select-none pointer-events-none">
                <FileText className="w-64 h-64 rotate-12" />
              </div>
              
              <div className="max-w-2xl mx-auto space-y-16 relative">
                <div className="text-center border-b border-brand-border pb-12">
                  <span className="inline-block px-4 py-1.5 bg-brand-bg text-brand-accent rounded-full text-[10px] font-extrabold uppercase tracking-widest mb-6">
                    Technical Onboarding v2.0
                  </span>
                  <h1 className="text-5xl font-extrabold tracking-tight text-brand-primary mb-4">
                    {selectedOrg?.organisation || "Select Assessment Pack"}
                  </h1>
                  <p className="text-sm font-medium text-brand-muted uppercase tracking-[0.2em]">DCM Assessment Pack v2.0</p>
                </div>

                <div className="grid grid-cols-2 gap-10">
                  <BriefField label="Executive Sponsor" value={selectedOrg?.staff_members?.split(',')[0] || "TBC"} />
                  <BriefField label="Financial Tier" value={selectedOrg?.revenue || "Assessment Pending"} />
                </div>

                <div className="space-y-12">
                  <BriefSection 
                    title="Strategic Summary" 
                    content={selectedOrg?.briefSummary || "No strategic summary generated yet. Navigate to 'Scoring' to add a project brief and distill objectives."}
                    isMarkdown
                  />
                  <BriefSection 
                    title="Project Objectives" 
                    content={selectedOrg ? `Define the core data-driven goals for ${selectedOrg.organisation}. What specific social impact does this project aim to catalyze through visualization or predictive modeling?` : "Define the core data-driven goals..."}
                  />
                  <BriefSection 
                    title="Inventory of Data Assets" 
                    content="Enumerate the available datasets, including source systems (CRM, ERP, Surveys), update frequency, and known schema limitations for initial exploratory analysis."
                  />
                  <BriefSection 
                    title="Stakeholder Impact Map" 
                    content="Identify the primary audience (Board, Field Staff, Donors) and the critical decisions this tool will empower. Mapping these prevents 'scope creep' during volunteer development."
                  />

                  {(consolidatedData?.dcmComment || consolidatedData?.dcmName) && (
                    <div className="pt-10 border-t-4 border-brand-accent bg-brand-bg/30 p-8 rounded-2xl">
                       <h4 className="text-[10px] font-black uppercase tracking-[0.3em] text-brand-accent mb-4">Executive Recommendation (DCM)</h4>
                       <div className="text-sm font-medium leading-relaxed text-brand-primary italic mb-6">
                         "{consolidatedData?.dcmComment || "No comment provided."}"
                       </div>
                       <div className="flex items-center gap-3">
                         <div className="w-8 h-8 rounded-full bg-brand-accent flex items-center justify-center text-white font-bold text-xs">
                           {consolidatedData?.dcmName?.[0] || "?"}
                         </div>
                         <div>
                            <p className="text-xs font-bold text-brand-primary">{consolidatedData?.dcmName || "Anonymous Advisor"}</p>
                            <p className="text-[9px] font-bold text-brand-muted uppercase tracking-widest">DCM Technical Consultant</p>
                         </div>
                       </div>
                    </div>
                  )}
                </div>

                <div className="pt-20 border-t border-brand-border flex justify-between items-center opacity-40">
                  <p className="text-[10px] font-bold uppercase tracking-widest italic">Changemaker Systems • Internal Confidential</p>
                  <div className="w-8 h-8 bg-brand-primary rounded flex items-center justify-center text-white text-[10px] font-black">Δ</div>
                </div>
              </div>
            </div>
          )}
        </div>
      </div>

      <AnimatePresence>
        {showRecommendation && selectedOrg && orgAssessment && (
            <RecommendationPack 
              lead={selectedOrg} 
              assessment={orgAssessment} 
              consolidatedData={consolidatedData}
              notify={notify}
              userProfile={userProfile}
              onClose={() => setShowRecommendation(false)} 
            />
        )}
      </AnimatePresence>
    </motion.div>
  );
}

function RecommendationPack({ lead, assessment, onClose, notify, consolidatedData, userProfile }: { lead: any, assessment: any, onClose: () => void, notify: (msg: string, type: 'success' | 'error') => void, consolidatedData?: any, userProfile?: any }) {
  const packRef = useRef<HTMLDivElement>(null);
  const [generating, setGenerating] = useState(false);
  const getVerificationIcon = (v: any) => v?.value ? "✅" : "❌";
  
  const dName = userProfile?.name || lead.dcmName || "Technical Consultant";
  
  const data = consolidatedData || {
    briefSummary: lead.briefSummary,
    revenue: lead.revenue,
    ein: lead.ein,
    charity_navigator_rating: lead.charity_navigator_rating,
    propublica_grants: lead.propublica_grants,
    linkedin_overview: lead.linkedin_overview || lead.linkedin_activity,
    staff_linkedin_summary: (lead.staff_linkedin_summary && lead.staff_linkedin_summary.length > 20) ? lead.staff_linkedin_summary : lead.staff_members,
    dcmComment: lead.dcmComment,
    dcmName: dName
  };

  const revenueNum = parseRevenueValue(lead.revenue || "");
  const adminFee = revenueNum > 1000000 ? 1000 : 500;

  const dataAudit = (assessment?.dataAssessment || lead?.dataAssessment) as DataQualityAssessment | null;
  const score = assessment?.averageScore || 0;
  const status = assessment?.status || "pending";
  const dateStr = assessment ? (assessment.updatedAt || assessment.createdAt) : new Date().toISOString();
  const dateFormatted = new Date(dateStr).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
  
  return (
    <motion.div 
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      className="fixed inset-0 z-50 bg-brand-primary/95 flex items-center justify-center p-6 backdrop-blur-sm print-bg overflow-y-auto"
    >
      <motion.div 
        initial={{ scale: 0.9, y: 20 }}
        animate={{ scale: 1, y: 0 }}
        className="bg-white rounded-3xl w-full max-w-5xl my-auto card-shadow relative print-area"
        ref={packRef}
      >
        <button 
          onClick={onClose}
          className="absolute top-6 right-6 p-2 hover:bg-brand-bg rounded-lg transition-all no-print"
        >
          <X className="w-6 h-6 text-brand-primary" />
        </button>

        <div className="p-10 md:p-16">
          <div className="mb-16 text-center">
            <div className="inline-flex items-center gap-2 px-4 py-1.5 bg-brand-accent/10 text-brand-accent rounded-full text-[10px] font-black uppercase tracking-[0.2em] mb-10">
              <ShieldCheck className="w-3.5 h-3.5" /> Intelligence Strategy Pack • Confidential
            </div>
            <h1 className="text-3xl md:text-4xl font-black tracking-tight text-brand-primary mb-6">{lead.organisation}</h1>
            <p className="text-sm md:text-base font-medium text-brand-muted max-w-2xl mx-auto px-4 italic underline decoration-brand-accent/20 decoration-2 underline-offset-8">
              "Strategic Roadmap for AI Transformation and Data Partnership Readiness."
            </p>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-12 gap-12 mb-20">
            <div className="lg:col-span-9 space-y-16">
              <section>
                <SectionHeader label="Strategic Summary & Project Objectives" className="border-l-4 border-brand-accent pl-4" />
                <div className="bg-brand-bg/50 p-8 rounded-2xl border border-brand-border text-[13px] md:text-[14px] font-medium text-brand-primary leading-relaxed italic markdown-content shadow-inner">
                  <Markdown>
                    {String(data.briefSummary || "No strategic objective summary recorded.")}
                  </Markdown>
                </div>
              </section>

              <section className="grid grid-cols-1 md:grid-cols-2 gap-10">
                <div className="space-y-6">
                  <SectionHeader label="AI Financial Intelligence" />
                  <div className="space-y-4">
                    <div className="p-6 bg-brand-bg rounded-2xl border border-brand-border text-brand-primary shadow-sm">
                       <p className="text-[10px] font-bold text-brand-muted uppercase tracking-widest mb-3">Revenue & Entity Status</p>
                       <p className="text-xs font-medium leading-relaxed tabular-nums markdown-content">{lead.revenue}</p>
                       <p className="text-[10px] font-mono text-brand-muted mt-3 pt-3 border-t border-brand-border/50">EIN Verified: {lead.ein || "N/A"}</p>
                    </div>
                    {data.propublica_grants && (
                      <div className="p-6 bg-brand-bg/30 rounded-2xl border border-brand-outline text-[10px] font-medium leading-relaxed markdown-content text-brand-primary">
                        <Markdown>{String(data.propublica_grants)}</Markdown>
                      </div>
                    )}
                  </div>
                </div>

                <div className="space-y-6">
                  <SectionHeader label="Partner Reputation Ranking" />
                  <div className="h-full">
                    <div className="p-8 bg-brand-primary text-white rounded-2xl flex flex-col justify-center shadow-xl shadow-brand-primary/10">
                       <p className="text-[10px] font-bold uppercase tracking-widest text-brand-accent mb-4">Charity Navigator Audit</p>
                       <div className="text-[10px] markdown-content font-medium text-white/90">
                        <Markdown>{String(data.charity_navigator_rating || "Reputation Index: Pending Manual Review")}</Markdown>
                      </div>
                    </div>
                  </div>
                </div>
              </section>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-10">
                <div className="space-y-6">
                  <SectionHeader label="1. Foundation Verification" />
                  <div className="bg-brand-bg p-6 rounded-2xl border border-brand-border space-y-4 border-l-4 border-emerald-500">
                    {[
                      { l: "Reputable & Traceable", v: assessment.verification?.reputable },
                      { l: "Mission Compatibility", v: assessment.verification?.fitsMission },
                      { l: "Non-Profit Verification", v: assessment.verification?.isNonProfit },
                      { l: "Delivery Window (2M)", v: assessment.verification?.available }
                    ].map((item, idx) => (
                      <div key={idx} className="flex justify-between items-center text-xs font-bold text-brand-primary">
                        <span className="text-brand-muted font-mono">{item.l}</span>
                        <span>{getVerificationIcon(item.v)}</span>
                      </div>
                    ))}
                  </div>
                </div>

                <div className="space-y-6">
                  <SectionHeader label="2. Dynamic Scorecard" />
                  <div className="bg-brand-bg p-6 rounded-2xl border border-brand-border space-y-5">
                    <ScoreBar label="Data Maturity" value={assessment.validation?.dataQuality?.value || 0} />
                    <ScoreBar label="Mission Alignment" value={assessment.validation?.missionAlignment?.value || 0} />
                    <ScoreBar label="Funding Readiness" value={assessment.validation?.fundsAvailable?.value || 0} />
                  </div>
                </div>
              </div>

              {dataAudit && (
                <section>
                  <SectionHeader label="3. AI Technical Data Audit Intelligence" className="text-brand-accent" />
                  <div className="bg-brand-primary rounded-2xl p-8 text-white space-y-8 shadow-2xl shadow-brand-accent/10">
                    <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 pb-6 border-b border-white/10 text-brand-accent">
                      <div>
                        <p className="text-[10px] font-black uppercase tracking-widest text-white/40">Data Visualisation Score</p>
                        <p className="text-4xl font-black">{dataAudit.score}<span className="text-xl opacity-40">/10</span></p>
                      </div>
                      <div className="text-right">
                        <p className="text-[10px] font-black uppercase tracking-widest text-white/40">Audit Verdict</p>
                        <p className="text-lg font-bold italic">"{dataAudit.verdict}"</p>
                      </div>
                    </div>
                    
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-10">
                      <div className="space-y-4">
                        <h6 className="text-[10px] font-black uppercase tracking-widest text-emerald-400">Inventory Strengths</h6>
                        <ul className="space-y-2">
                          {dataAudit.strengths.map((s, i) => (
                            <li key={i} className="text-xs font-medium flex gap-3 text-white/80">
                              <span className="text-emerald-400 font-bold">✓</span> {s}
                            </li>
                          ))}
                        </ul>
                      </div>
                      <div className="space-y-4">
                        <h6 className="text-[10px] font-black uppercase tracking-widest text-amber-400">Growth Roadblocks</h6>
                        <ul className="space-y-2">
                          {dataAudit.weaknesses.map((w, i) => (
                            <li key={i} className="text-xs font-medium flex gap-3 text-white/80">
                              <span className="text-amber-400 font-bold">!</span> {w}
                            </li>
                          ))}
                        </ul>
                      </div>
                    </div>

                    <div className="pt-6 border-t border-white/10">
                      <h6 className="text-[10px] font-black uppercase tracking-widest text-brand-accent mb-3">Technical Strategic Roadmap</h6>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        {dataAudit.recommendations.map((r, i) => (
                          <div key={i} className="bg-white/5 border border-white/10 p-3 rounded-lg text-[11px] font-semibold text-white/90">
                             {r}
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                </section>
              )}

              <section className="grid grid-cols-1 md:grid-cols-2 gap-12">
                 <div className="space-y-6">
                   <SectionHeader label="Digital Reach (LinkedIn Audit)" />
                   <div className="bg-brand-bg/50 p-6 rounded-2xl border border-brand-border text-xs leading-relaxed font-medium markdown-content text-brand-primary max-h-60 overflow-y-auto custom-scrollbar">
                     <Markdown>{String(data.linkedin_overview || "No digital audit performed.")}</Markdown>
                   </div>
                 </div>
                 <div className="space-y-6">
                   <SectionHeader label="Leadership Blueprint" />
                   <div className="bg-brand-bg/50 p-6 rounded-2xl border border-brand-border text-xs leading-relaxed font-medium markdown-content text-brand-primary max-h-60 overflow-y-auto custom-scrollbar">
                     <Markdown>{String(data.staff_linkedin_summary || "Manual leadership review advised.")}</Markdown>
                   </div>
                 </div>
              </section>

              {(data.dcmComment || data.dcmName) && (
                <section className="pt-10 border-t-4 border-brand-accent bg-brand-bg/30 p-10 rounded-3xl relative overflow-hidden">
                   <div className="absolute top-0 right-0 p-8 opacity-5 pointer-events-none">
                     <ShieldCheck className="w-40 h-40 text-brand-accent" />
                   </div>
                   <div className="relative">
                     <h4 className="text-[11px] font-black uppercase tracking-[0.4em] text-brand-accent mb-8 flex items-center gap-2">
                       <Target className="w-4 h-4" /> Executive Recommendation (DCM)
                     </h4>
                     <div className="text-sm font-extrabold md:text-base leading-relaxed text-brand-primary italic mb-10 border-l-4 border-brand-accent pl-8 py-2">
                       "{data.dcmComment || "Strategic induction recommended with standard monitoring."}"
                     </div>
                     <div className="flex items-center gap-4">
                       <div className="w-12 h-12 rounded-2xl bg-brand-accent flex items-center justify-center text-white font-black text-lg shadow-lg shadow-brand-accent/20">
                         {data.dcmName?.[0] || "Δ"}
                       </div>
                       <div>
                          <p className="text-sm font-black text-brand-primary leading-tight">{data.dcmName || "Technical Advisor"}</p>
                          <p className="text-[10px] font-bold text-brand-muted uppercase tracking-widest mt-1">DCM Technical Strategy Consultant</p>
                       </div>
                     </div>
                   </div>
                </section>
              )}
            </div>

            <div className="lg:col-span-3 space-y-12">
              <div className="sticky top-10 space-y-8">
                <div>
                  <SectionHeader label="Quick Metrics" />
                  <div className="space-y-6 bg-brand-bg p-6 rounded-2xl border border-brand-border shadow-sm">
                    <div className="flex gap-4 items-start">
                      <div className="p-2 bg-white text-brand-accent rounded-lg shadow-sm border border-brand-border/50"><DollarSign className="w-5 h-5" /></div>
                      <div>
                        <p className="text-[10px] font-bold text-brand-muted uppercase tracking-widest mb-1">Scale</p>
                        <p className="text-[10px] font-semibold text-brand-primary mt-1">{lead.revenue}</p>
                      </div>
                    </div>
                    <div className="flex gap-4 items-start">
                      <div className="p-2 bg-white text-brand-accent rounded-lg shadow-sm border border-brand-border/50"><ShieldCheck className="w-5 h-5" /></div>
                      <div>
                        <p className="text-[10px] font-bold text-brand-muted uppercase tracking-widest mb-1">Key Leadership</p>
                        <p className="text-[11px] font-extrabold text-brand-primary leading-tight line-clamp-3 italic">
                          {data.staff && data.staff !== "Manual leadership review advised." ? data.staff : "Identified via Research"}
                        </p>
                      </div>
                    </div>
                  </div>
                </div>

                <div className={cn(
                  "p-8 rounded-2xl border-2 text-center shadow-lg",
                  status === "approved" ? "bg-emerald-50/50 border-emerald-200" : 
                  status === "pending" ? "bg-amber-50/50 border-amber-200" :
                  status === "not suitable" ? "bg-rose-50/50 border-rose-200" : "bg-blue-50/50 border-blue-200"
                )}>
                  <div className="w-16 h-16 mx-auto bg-white rounded-full flex items-center justify-center text-3xl mb-5 shadow-sm border border-brand-border">
                    {status === "approved" ? "✅" : 
                     status === "pending" ? "⏳" :
                     status === "not suitable" ? "⛔" : "ℹ️"}
                  </div>
                  <h4 className={cn(
                     "text-lg font-black mb-1 uppercase tracking-tight",
                     status === "approved" ? "text-emerald-900" : 
                     (status === "pending" || (lead.leadStatus || "").toLowerCase().includes("under assessment")) ? "text-amber-900" :
                     status === "not suitable" ? "text-rose-900" : "text-blue-900"
                  )}>
                    {status === "approved" ? "ASSESSMENT APPROVED" : 
                     (status === "pending" || (lead.leadStatus || "").toLowerCase().includes("under assessment")) ? "UNDER ASSESSMENT" :
                     status === "not suitable" ? "NOT SUITABLE" : "MORE INFO"}
                  </h4>
                  <p className="text-[8px] font-bold text-brand-muted uppercase tracking-widest mb-4">
                    Assessed by {data.dcmName} on {dateFormatted}
                  </p>
                  <p className="text-[10px] font-bold text-brand-muted uppercase tracking-widest mb-6 border-b border-brand-border pb-4">
                    Scoring Index: {score}/4.0
                  </p>
                  
                  <div className="space-y-1">
                     <p className="text-[9px] font-bold text-brand-muted uppercase tracking-widest">Admin Fee Baseline</p>
                     <p className="text-2xl font-black text-brand-primary italic">${adminFee.toLocaleString()}</p>
                     <p className="text-[8px] text-brand-muted uppercase font-bold tracking-tighter">(Based on Revenue Scale)</p>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div className="mb-20 pt-16 border-t border-brand-border">
            <SectionHeader label="4. Consolidation & Validation Audit" />
            <div className="grid grid-cols-1 md:grid-cols-2 gap-10">
              <div className="bg-brand-bg rounded-2xl p-8 border border-brand-border border-l-4 border-brand-primary shadow-sm">
                <h5 className="text-[11px] font-black uppercase tracking-wider text-brand-muted mb-8">Strategic Intelligence Radar</h5>
                <div className="space-y-4">
                  {[
                    { label: "Calculated Quality Index", val: `${score} / 4.0`, color: "text-brand-accent" },
                    { label: "Technical Readiness Audit", val: dataAudit ? "COMPLETED" : "MANUAL REVIEW", color: dataAudit ? "text-emerald-600" : "text-amber-600" },
                    { label: "Partnership Outcome", val: status === 'approved' ? 'READY' : 'PENDING', color: status === 'approved' ? 'text-emerald-600' : 'text-brand-muted' }
                  ].map((item, i) => (
                    <div key={i} className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-1 border-b border-brand-border pb-3 last:border-0 last:pb-0">
                      <span className="text-[11px] font-bold text-brand-muted uppercase tracking-tight">{item.label}</span>
                      <span className={cn("text-xs font-extrabold", item.color)}>{item.val}</span>
                    </div>
                  ))}
                </div>
              </div>
              <div className="bg-brand-primary text-white rounded-2xl p-8 relative overflow-hidden flex flex-col justify-center shadow-xl">
                 <div className="absolute -top-4 -right-4 opacity-5 pointer-events-none">
                   <ShieldCheck className="w-32 h-32" />
                 </div>
                 <h5 className="text-[10px] font-black uppercase tracking-widest text-emerald-400 mb-4 flex items-center gap-2">
                   <Shield className="w-3 h-3" /> Audit Assurance
                 </h5>
                 <p className="text-xs font-medium leading-relaxed text-white/90 italic">
                   "This pack consolidates multi-agent deep AI research and human-verified scoring. All mission-critical flags have been assessed and verified by {data.dcmName} as of {dateFormatted}."
                 </p>
              </div>
            </div>
          </div>

          <div className="pt-10 border-t-2 border-brand-border flex flex-col md:flex-row justify-between items-center gap-8 no-print">
            <div className="flex flex-col gap-2 flex-1">
               <p className="text-[10px] font-bold text-brand-muted uppercase tracking-widest italic">
                 * Save this PDF then upload to the Lead record in the DCM Hub as an artifact.
               </p>
               <button 
                disabled={generating}
                onClick={async () => {
                  if (!packRef.current) return;
                  setGenerating(true);
                  try {
                  const content = packRef.current;
                  
                  // Use toJpeg with 0.85 quality for a much smaller file size than PNG
                  const dataUrl = await toJpeg(content, {
                    quality: 0.85,
                    pixelRatio: 1.5, // Reduced from 2 to further save space
                    style: {
                      background: '#ffffff',
                      width: '1024px',
                      margin: '0',
                      padding: '0'
                    },
                    filter: (node) => {
                      if (node instanceof HTMLElement) {
                        return !node.classList.contains('no-print');
                      }
                      return true;
                    }
                  });

                  const pdf = new jsPDF({
                    orientation: 'p',
                    unit: 'mm',
                    format: 'a4',
                    compress: true // Enable jsPDF internal compression
                  });
                  const pdfWidth = 210; // A4 width in mm
                const pdfHeight = 297; // A4 height in mm
                
                const imgProps = pdf.getImageProperties(dataUrl);
                const imgHeight = (imgProps.height * pdfWidth) / imgProps.width;
                
                let heightLeft = imgHeight;
                let position = 0;

                // Add the image to PDF with multi-page support
                while (heightLeft > 0) {
                  // Add image with negative vertical offset to show the correct slice
                  pdf.addImage(dataUrl, 'PNG', 0, position, pdfWidth, imgHeight);
                  heightLeft -= pdfHeight;
                  position -= pdfHeight;
                  
                  // Only add a new page if we have significant height remaining (> 1mm)
                  if (heightLeft > 1) {
                    pdf.addPage();
                  }
                }
                
                pdf.save(`${String(lead.organisation || "Organization").replace(/\s+/g, '_')}_Assessment_Pack.pdf`);
                  
                  await setDoc(doc(db, "leads", lead.id), { leadStatus: "Assessed", updatedAt: new Date().toISOString() }, { merge: true });
                  notify?.("Assessment Pack generated and downloaded. Status updated to 'Assessed'.", "success");
                } catch (e) {
                  console.error("Failed to generate PDF:", e);
                  notify?.("Failed to generate PDF. Please try again or use standard print.", "error");
                } finally {
                  setGenerating(false);
                }
              }}
              className="w-full md:w-auto px-10 py-5 bg-brand-primary text-white rounded-2xl font-black text-xs uppercase tracking-widest flex items-center justify-center gap-3 hover:bg-brand-primary/95 hover:scale-[1.02] active:scale-[0.98] transition-all shadow-2xl shadow-brand-primary/20 disabled:opacity-50"
            >
              {generating ? <Loader2 className="w-5 h-5 animate-spin" /> : <Download className="w-5 h-5 text-brand-accent" />}
              {generating ? "GENERATING PDF..." : "GENERATE ASSESSMENT PACK PDF"}
            </button>
          </div>
            <div className="text-center md:text-right">
               <p className="text-[10px] font-black text-brand-primary uppercase tracking-[0.3em] mb-1">Changemaker Systems AI Core</p>
               <p className="text-[9px] font-bold text-brand-muted/60 font-mono italic inline-flex items-center gap-1"><Shield className="w-3 h-3 text-brand-accent" /> AUTHENTICATED INTEL REPORT V2</p>
            </div>
          </div>
        </div>
      </motion.div>
    </motion.div>
  );
}

// --- Helper UI Components ---

function StatsCard({ icon: Icon, label, value, subLabel, onClick, isActive }: { icon: any, label: string, value: string, subLabel?: string, onClick?: () => void, isActive?: boolean }) {
  return (
    <div 
      onClick={onClick}
      className={cn(
        "bg-white p-6 rounded-xl border transition-all card-shadow group",
        onClick ? "cursor-pointer hover:border-brand-accent/50" : "hover:border-brand-accent/30",
        isActive ? "border-brand-accent bg-brand-accent/5 ring-1 ring-brand-accent/20" : "border-brand-border"
      )}
    >
      <div className="flex items-start justify-between mb-4">
        <div className={cn(
          "p-2 rounded-lg transition-all",
          isActive ? "bg-brand-accent text-white" : "bg-brand-bg text-brand-accent group-hover:bg-brand-accent group-hover:text-white"
        )}>
          <Icon className="w-5 h-5" />
        </div>
        {subLabel && (
          <span className={cn(
            "text-[9px] font-extrabold uppercase tracking-widest px-2 py-0.5 rounded-sm",
            isActive ? "bg-white text-brand-accent" : "bg-brand-success/10 text-brand-success"
          )}>
            {subLabel}
          </span>
        )}
      </div>
      <h4 className={cn(
        "text-[10px] font-bold uppercase tracking-widest mb-1",
        isActive ? "text-brand-accent" : "text-brand-muted"
      )}>{label}</h4>
      <p className={cn(
        "text-2xl font-extrabold",
        isActive ? "text-brand-accent" : "text-brand-primary"
      )}>{value}</p>
    </div>
  );
}

function RubricHeader({ label, description }: { label: string, description: string }) {
  return (
    <div className="flex items-end gap-4 mb-4">
      <h4 className="text-lg font-black text-brand-primary leading-none uppercase tracking-tight">{label}</h4>
      <p className="text-[10px] font-bold text-brand-muted uppercase tracking-widest leading-none mb-0.5">{description}</p>
    </div>
  );
}

function RubricRow({ 
  label, 
  type, 
  value, 
  comment, 
  onChange, 
  onCommentChange, 
  hint 
}: { 
  label: string, 
  type: 'boolean' | 'score', 
  value: any, 
  comment: string, 
  onChange: (v: any) => void, 
  onCommentChange: (c: string) => void,
  hint?: string
}) {
  return (
    <div className="p-8 bg-white hover:bg-brand-bg/20 transition-all flex flex-col lg:flex-row gap-8 items-start">
      <div className="flex-1 w-full lg:w-auto">
        <p className="text-sm font-bold text-brand-primary leading-snug mb-2">{label}</p>
        {hint && (
          <div className="inline-block bg-amber-50 border border-amber-100 rounded px-2 py-1 mt-1">
             <p className="text-[9px] font-bold text-amber-700 italic leading-tight">{hint}</p>
          </div>
        )}
      </div>
      
      <div className="w-full lg:w-48 shrink-0">
        {type === 'boolean' ? (
          <div className="flex gap-1.5 bg-brand-bg p-1 rounded-lg border border-brand-border/50">
            <button 
              onClick={() => onChange(true)}
              className={cn("flex-1 py-2 rounded-md text-[10px] font-black uppercase tracking-widest transition-all", value === true ? "bg-white text-emerald-600 shadow-sm" : "text-brand-muted hover:text-brand-text")}
            >
              YES
            </button>
            <button 
              onClick={() => onChange(false)}
              className={cn("flex-1 py-2 rounded-md text-[10px] font-black uppercase tracking-widest transition-all", value === false ? "bg-white text-red-500 shadow-sm" : "text-brand-muted hover:text-brand-text")}
            >
              NO
            </button>
          </div>
        ) : (
          <div className="space-y-3">
             <div className="flex justify-between items-end">
               <span className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Score 1-4</span>
               <span className="text-lg font-black text-brand-accent leading-none">{value}</span>
             </div>
             <div className="flex gap-1.5">
               {[1, 2, 3, 4].map(n => (
                 <button 
                   key={n}
                   onClick={() => onChange(n)}
                   className={cn(
                     "flex-1 h-2 rounded-full transition-all",
                     n <= value ? "bg-brand-accent" : "bg-brand-border"
                   )}
                 />
               ))}
             </div>
          </div>
        )}
      </div>

      <div className="w-full lg:w-96 shrink-0">
         <h5 className="text-[10px] font-bold text-brand-muted uppercase tracking-widest mb-2">Comments</h5>
         <textarea 
            value={comment}
            onChange={(e) => onCommentChange(e.target.value)}
            placeholder="Add reasoning..."
            className="w-full h-20 bg-brand-bg/30 border border-brand-border rounded-lg p-3 text-xs font-semibold placeholder:text-brand-muted/50 outline-none focus:ring-2 focus:ring-brand-accent/20 transition-all resize-none"
         />
      </div>
    </div>
  );
}

function SectionHeader({ label, className }: { label: string, className?: string }) {
  return <h4 className={cn("text-[10px] font-bold uppercase tracking-widest text-brand-muted mb-4 flex items-center gap-2 after:content-[''] after:flex-1 after:h-[1px] after:bg-brand-border", className)}>{label}</h4>;
}

function ScoreBar({ label, value }: { label: string, value: number }) {
  return (
    <div className="space-y-1.5">
      <div className="flex justify-between text-[10px] font-bold uppercase tracking-widest text-brand-muted mb-1">
        <span>{label}</span>
        <span>{value} / 4</span>
      </div>
      <div className="h-1.5 bg-brand-border rounded-full overflow-hidden">
        <div 
          className="h-full bg-brand-accent transition-all duration-1000" 
          style={{ width: `${(value / 4) * 100}%` }}
        />
      </div>
    </div>
  );
}

function BriefField({ label, value }: { label: string, value: string }) {
  return (
    <div className="bg-brand-bg/50 p-4 rounded-xl border border-brand-border/50">
      <h5 className="text-[9px] font-extrabold uppercase tracking-[0.2em] text-brand-muted mb-2">{label}</h5>
      <p className="font-bold text-brand-primary">{value}</p>
    </div>
  );
}

function BriefSection({ title, content, isMarkdown }: { title: string, content: string, isMarkdown?: boolean }) {
  return (
    <div className="space-y-4">
      <h4 className="flex items-center gap-3 text-[11px] font-extrabold uppercase tracking-[0.25em] text-brand-primary">
        <span className="w-1.5 h-1.5 bg-brand-accent rounded-full animate-pulse"></span>
        {title}
      </h4>
      <div className="text-sm font-medium leading-relaxed text-brand-muted pl-5 border-l-2 border-brand-border/30 markdown-content">
        {isMarkdown ? <Markdown>{String(content || "")}</Markdown> : <p>{content}</p>}
      </div>
    </div>
  );
}

function LegendItem({ label, color }: { label: string, color: string }) {
  return (
    <div className="flex items-center gap-2">
      <div className="w-2.5 h-2.5 rounded-full" style={{ backgroundColor: color }}></div>
      <span className="text-[10px] font-bold uppercase tracking-widest text-brand-muted">{label}</span>
    </div>
  );
}

function LeadDetailItem({ icon: Icon, label, value, isBadge, isLink, onSave, options, type = "text", disabled = false }: { 
  icon: any, 
  label: string, 
  value?: string, 
  isBadge?: boolean,
  isLink?: boolean,
  onSave?: (newValue: string) => void,
  options?: string[],
  type?: string,
  disabled?: boolean
}) {
  const [editing, setEditing] = useState(false);
  const [localValue, setLocalValue] = useState(value || "");
  const isEmpty = (!value || value === "Empty") && !editing;

  useEffect(() => {
    setLocalValue(value || "");
  }, [value]);

  const handleManualSave = () => {
    if (onSave) {
      onSave(localValue);
    }
    setEditing(false);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey && type !== "textarea") {
      handleManualSave();
    }
    if (e.key === 'Escape') {
      setLocalValue(value || "");
      setEditing(false);
    }
  };
  
  return (
    <div className="flex items-start gap-4 group">
      <div className="w-8 flex justify-center pt-0.5 text-white/30 group-hover:text-brand-accent transition-colors">
        <Icon className="w-4 h-4" />
      </div>
      <div className="flex-1 border-b border-white/5 pb-4">
        <div className="flex justify-between items-center mb-1">
          <span className="text-[10px] font-black uppercase tracking-[0.2em] text-white/20">{label}</span>
          {editing ? (
            <div className="flex gap-2">
              <button 
                onClick={() => { setLocalValue(value || ""); setEditing(false); }}
                className="text-[9px] font-bold text-white/40 hover:text-white transition-all uppercase tracking-widest"
              >
                Cancel
              </button>
              <button 
                onClick={handleManualSave}
                className="text-[9px] font-bold text-brand-accent hover:text-brand-accent/80 transition-all uppercase tracking-widest"
              >
                Save
              </button>
            </div>
          ) : onSave && !disabled && (
            <button 
              onClick={() => setEditing(true)}
              className="text-[9px] font-bold text-white/20 hover:text-brand-accent opacity-0 group-hover:opacity-100 transition-all uppercase tracking-widest"
            >
              Edit
            </button>
          )}
        </div>
        {editing ? (
          <div className="flex gap-2 items-end pt-1">
            {options ? (
              <select
                autoFocus
                className="flex-1 bg-white/10 border border-white/20 rounded-lg px-3 py-2 text-sm font-medium text-white outline-none focus:ring-1 focus:ring-brand-accent transition-all appearance-none cursor-pointer"
                value={localValue}
                onChange={(e) => setLocalValue(e.target.value)}
              >
                {options.map(opt => (
                  <option key={opt} value={opt} className="bg-[#1A1A1A] text-white">{opt}</option>
                ))}
              </select>
            ) : type === "date" ? (
              <input 
                type="date"
                autoFocus
                className="flex-1 bg-white/5 border border-white/10 rounded-lg px-3 py-2 text-sm font-medium text-white outline-none focus:ring-1 focus:ring-brand-accent transition-all"
                value={localValue}
                onChange={(e) => setLocalValue(e.target.value)}
                onKeyDown={handleKeyDown}
              />
            ) : (
              <textarea
                autoFocus
                className="flex-1 bg-white/5 border border-white/10 rounded-lg px-3 py-2 text-sm font-medium text-white outline-none focus:ring-1 focus:ring-brand-accent transition-all resize-none min-h-[80px]"
                value={localValue}
                onChange={(e) => setLocalValue(e.target.value)}
                onKeyDown={handleKeyDown}
              />
            )}
          </div>
        ) : isEmpty ? (
          <span className={cn("text-sm font-medium text-white/20 italic", !disabled && "cursor-pointer")} onClick={() => onSave && !disabled && setEditing(true)}>Empty</span>
        ) : isBadge ? (
          <span className={cn(
            "inline-block text-[9px] font-black uppercase tracking-widest px-2.5 py-1 rounded",
            !disabled && "cursor-pointer",
            label.toLowerCase().includes("status") ? "bg-brand-accent/20 text-brand-accent" : "bg-white/10 text-white/70"
          )} onClick={() => onSave && !disabled && setEditing(true)}>
            {value}
          </span>
        ) : isLink ? (
          <div className="flex items-center gap-2">
            <a href={value?.startsWith('http') ? value : `https://${value}`} target="_blank" rel="noreferrer" className="text-sm font-medium text-brand-accent hover:underline break-all">
              {value}
            </a>
            {!disabled && (
              <button onClick={() => onSave && setEditing(true)} className="text-white/20 hover:text-white opacity-0 group-hover:opacity-100 transition-all">
                <FastForward className="w-3 h-3 rotate-90" />
              </button>
            )}
          </div>
        ) : (
          <p className={cn("text-sm font-medium text-white/80 leading-relaxed break-words", !disabled && "cursor-pointer")} onClick={() => onSave && !disabled && setEditing(true)}>{value}</p>
        )}
      </div>
    </div>
  );
}

// DCM Leads (CRM) Section
function EditableTableCell({ value, onSave, type = "text", options, format, className, disabled = false }: { 
  value: string, 
  onSave: (val: string) => void, 
  type?: "text" | "date" | "select",
  options?: string[],
  format?: string,
  className?: string,
  disabled?: boolean
}) {
  const [isEditing, setIsEditing] = useState(false);
  const [localValue, setLocalValue] = useState(value);

  useEffect(() => {
    setLocalValue(value);
  }, [value]);

  const handleManualSave = () => {
    setIsEditing(false);
    if (localValue !== value) {
      onSave(localValue);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      handleManualSave();
    }
    if (e.key === 'Escape') {
      setLocalValue(value);
      setIsEditing(false);
    }
  };

  if (isEditing) {
    const inputClasses = "bg-white border border-brand-accent rounded px-2 py-0.5 text-xs font-medium outline-none shadow-sm focus:ring-2 focus:ring-brand-accent/20 transition-all flex-1";
    
    return (
      <div className="flex items-center gap-1.5 w-full" onClick={(e) => e.stopPropagation()}>
        {type === "select" && options ? (
          <select 
            autoFocus
            className={inputClasses}
            value={localValue || ""}
            onChange={(e) => setLocalValue(e.target.value)}
          >
            <option value="">-</option>
            {options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
          </select>
        ) : type === "date" ? (
          <input 
            type="date"
            autoFocus
            className={inputClasses}
            value={localValue || ""}
            onChange={(e) => setLocalValue(e.target.value)}
            onKeyDown={handleKeyDown}
          />
        ) : (
          <input 
            autoFocus
            className={inputClasses}
            value={localValue || ""}
            onChange={(e) => setLocalValue(e.target.value)}
            onKeyDown={handleKeyDown}
          />
        )}
        <div className="flex items-center gap-1 shrink-0">
          <button 
            onClick={handleManualSave}
            className="p-1 rounded bg-brand-accent text-white hover:bg-brand-accent/90 shadow-sm transition-all"
          >
            <Check className="w-3 h-3" />
          </button>
          <button 
            onClick={() => { setLocalValue(value); setIsEditing(false); }}
            className="p-1 rounded bg-slate-100 text-slate-500 hover:bg-slate-200 transition-all"
          >
            <X className="w-3 h-3" />
          </button>
        </div>
      </div>
    );
  }

  return (
    <div 
      onDoubleClick={(e) => {
        if (disabled) return;
        e.stopPropagation();
        setIsEditing(true);
      }}
      className={cn(
        "cursor-text min-h-[20px] rounded px-1 transition-all", 
        disabled ? "cursor-default" : "hover:bg-brand-bg/50",
        className
      )}
      title={disabled ? "" : "Double click to edit"}
    >
      {type === "date" ? formatDate(value, format) : (value || '-')}
    </div>
  );
}

function LeadCardContent({ 
  lead, 
  isAdminOrBoard,
  canModifyStatus,
  canAddLead,
  canDeleteLead,
  setViewingLead, 
  handleDuplicate, 
  handleDelete, 
  onSelectResearch,
  boardStatuses, 
  notify, 
  db, 
  handleFirestoreError 
}: any) {
  return (
    <div className="bg-white p-4 rounded-xl border border-brand-border card-shadow cursor-default hover:border-brand-accent/30 transition-all group w-full">
      <div className="flex justify-between items-start mb-2">
        <h5 className="text-sm font-black text-brand-primary line-clamp-2 cursor-pointer hover:text-brand-accent transition-colors flex items-center gap-2" onClick={(e) => { e.stopPropagation(); setViewingLead(lead); }}>
          {lead.organisation}
          {lead.leadStatus === "Assessed" && (
            lead.approved ? (
              <div className="p-0.5 bg-emerald-500 rounded flex items-center justify-center shadow-[0_0_8px_rgba(16,185,129,0.4)]" title="Approved for Future Project">
                <CheckCircle2 className="w-2.5 h-2.5 text-white" />
              </div>
            ) : (
              <div className="p-0.5 bg-rose-500 rounded flex items-center justify-center shadow-[0_0_8px_rgba(244,63,94,0.4)]" title="Needs Board Approval">
                <ShieldAlert className="w-2.5 h-2.5 text-white" />
              </div>
            )
          )}
          {lead.isVerified && <ShieldCheck className="w-3.5 h-3.5 text-emerald-500 shrink-0" />}
          {lead.isPaidEngagement && <DollarSign className="w-3.5 h-3.5 text-brand-accent shrink-0" />}
        </h5>
        <div className="hidden group-hover:flex items-center gap-1">
          <button 
            onClick={(e) => { e.stopPropagation(); onSelectResearch?.(lead.id); }}
            className="p-1.5 text-brand-muted hover:text-brand-accent hover:bg-brand-bg rounded-md transition-colors"
            title="Deep Intelligence Research"
          >
            <Sparkles className="w-3.5 h-3.5" />
          </button>
          {canAddLead && (
            <button 
              onClick={(e) => { e.stopPropagation(); handleDuplicate(lead); }}
              className="p-1.5 text-brand-muted hover:text-brand-accent hover:bg-brand-bg rounded-md transition-colors"
              title="Duplicate Lead context (Admin)"
            >
              <Copy className="w-3.5 h-3.5" />
            </button>
          )}
          <button 
            onClick={(e) => { e.stopPropagation(); setViewingLead(lead); }}
            className="p-1.5 text-brand-muted hover:text-brand-accent hover:bg-brand-bg rounded-md transition-colors"
            title="Edit Lead Details"
          >
            <Edit3 className="w-3.5 h-3.5" />
          </button>
          {canDeleteLead && (
            <button 
              onClick={(e) => { e.stopPropagation(); handleDelete(lead.id); }}
              className="p-1.5 text-brand-muted hover:text-red-500 hover:bg-red-50 rounded-md transition-colors"
              title="Delete Lead"
            >
              <Trash2 className="w-3.5 h-3.5" />
            </button>
          )}
        </div>
      </div>
      <p className="text-[11px] font-medium text-brand-muted line-clamp-3 mb-4 leading-relaxed">
        {lead.activity || "No description provided."}
      </p>

      {/* Intelligence Badges */}
      <div className="flex flex-wrap gap-2 mb-4">
        {lead.briefSummary && (
          <div className="flex items-center gap-1.5 px-2.5 py-1 bg-brand-bg border border-brand-border rounded-lg shadow-sm" title="AI Assessment Complete">
            <Sparkles className="w-2.5 h-2.5 text-brand-accent animate-pulse" />
            <span className="text-[9px] font-black text-brand-primary uppercase tracking-widest">Research AI</span>
          </div>
        )}
        {lead.linkedin_overview && (
          <div className="flex items-center gap-1.5 px-2.5 py-1 bg-brand-bg border border-brand-border rounded-lg shadow-sm" title="LinkedIn Audit Complete">
            <Linkedin className="w-2.5 h-2.5 text-blue-500" />
            <span className="text-[9px] font-black text-brand-primary uppercase tracking-widest">Audit</span>
          </div>
        )}
      </div>

      {lead.needsFurtherAssessment && (
        <div className="mb-4 flex items-center gap-2 bg-amber-500/10 border border-amber-500/20 rounded-lg px-2 py-1.5 ring-1 ring-amber-500/20 shadow-sm shadow-amber-500/5">
          <AlertCircle className="w-2.5 h-2.5 text-amber-500 shrink-0" />
          <div className="flex flex-col">
            <span className="text-[8px] font-black uppercase tracking-widest text-amber-600">Needs Further Assessment</span>
          </div>
        </div>
      )}

      {(lead.leadStatus || "").toLowerCase() === "assessed" && !lead.approved && !lead.needsFurtherAssessment && (
        <div className="mb-4 flex items-center gap-2 bg-blue-500/10 border border-blue-500/20 rounded-lg px-2 py-1.5 ring-1 ring-blue-500/20 shadow-sm shadow-blue-500/5">
          <Clock className="w-2.5 h-2.5 text-blue-500 shrink-0" />
          <div className="flex flex-col">
            <span className="text-[8px] font-black uppercase tracking-widest text-blue-600">Waiting for Approval from DCM Board</span>
            {lead.assessedBy && (
              <span className="text-[8px] font-bold text-blue-600/60 uppercase tracking-widest mt-0.5">
                {lead.assessedBy.split('@')[0]} • {lead.assessedAt ? new Date(lead.assessedAt).toLocaleDateString() : ""}
              </span>
            )}
          </div>
        </div>
      )}

      {lead.assessedBy && lead.approved && !lead.needsFurtherAssessment && (
        <div className="mb-4 flex items-center gap-2 rounded-lg px-2 py-1.5 border bg-emerald-500/5 border-emerald-500/10">
          <CheckCircle2 className="w-2.5 h-2.5 text-emerald-500" />
          <div className="flex flex-col">
            <div className="text-[9px] font-black uppercase tracking-widest block text-emerald-600">
              Approved (Board)
            </div>
            <div className="text-[9px] font-bold text-brand-primary truncate block mt-0.5 opacity-80">
              {lead.approved 
                ? `${lead.approvedBy?.split('@')[0] || 'Admin'} • ${lead.approvedAt ? new Date(lead.approvedAt).toLocaleDateString() : '?'}`
                : `${(lead.assessedBy || "").split('@')[0]} • ${lead.assessedAt ? new Date(lead.assessedAt).toLocaleDateString() : "Pending"}`
              }
            </div>
          </div>
        </div>
      )}

      <div className="flex items-center justify-between pt-3 border-t border-brand-bg relative group/status">
        <div className="flex-1 overflow-hidden">
          <div className="flex items-center gap-1.5 min-w-0">
            <UserCircle className="w-2.5 h-2.5 text-brand-muted shrink-0" />
            <span className="text-[9px] font-bold text-brand-accent uppercase tracking-widest truncate">
              {lead.leadOwner || 'Unassigned'}
            </span>
          </div>
        </div>
        <div className="relative">
          {canModifyStatus ? (
            <>
              <select 
                className="absolute inset-0 opacity-0 cursor-pointer z-20"
                value={lead.leadStatus}
                onChange={(e) => {
                  const newStatus = e.target.value;
                  setDoc(doc(db, "leads", lead.id), { leadStatus: newStatus, updatedAt: new Date().toISOString() }, { merge: true })
                    .then(() => notify("Status updated", "success"))
                    .catch((err) => {
                      notify("Update failed", "error");
                      handleFirestoreError(err, OperationType.UPDATE, "leads");
                    });
                }}
              >
                {boardStatuses.map((s: string) => <option key={s} value={s}>{s}</option>)}
              </select>
              <div className="flex items-center gap-1 px-2 py-0.5 bg-brand-bg border border-brand-border rounded text-[8px] font-black text-brand-muted group-hover/status:border-brand-accent group-hover/status:text-brand-accent transition-all cursor-pointer">
                Status <ChevronDown className="w-2 h-2" />
              </div>
            </>
          ) : (
            <div className="flex items-center gap-1 px-2 py-0.5 bg-brand-bg/50 border border-brand-border rounded text-[8px] font-black text-brand-muted/40 cursor-not-allowed">
              {lead.leadStatus}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function DraggableLeadCard(props: any) {
  const { attributes, listeners, setNodeRef, transform, isDragging } = useDraggable({
    id: props.lead.id,
    data: { lead: props.lead }
  });

  const style = transform ? {
    transform: CSS.Translate.toString(transform),
  } : undefined;

  return (
    <div 
      ref={setNodeRef} 
      style={style} 
      {...attributes} 
      {...listeners}
      className={cn(isDragging && "opacity-20", "cursor-grab")}
    >
      <LeadCardContent {...props} />
    </div>
  );
}

function DroppableStatusColumn({ status, children }: { status: string, children: React.ReactNode }) {
  const { isOver, setNodeRef } = useDroppable({
    id: status,
    data: { status }
  });

  return (
    <div 
      ref={setNodeRef} 
      className={cn(
        "flex flex-col gap-3 min-h-[400px] transition-all p-1",
        isOver && "bg-brand-accent/5 ring-2 ring-brand-accent/20 rounded-xl"
      )}
    >
      {children}
    </div>
  );
}

function LeadsView({ leads, assessments, settings, userProfile, aiConfig, onShowAiConfig, handleAiError, notify, updateLeadsConfig, logLeadEvent, onSelectResearch, users = [] }: { 
  leads: any[], 
  assessments: any[],
  settings: any, 
  userProfile: any, 
  aiConfig: any,
  onShowAiConfig: () => void,
  handleAiError: (err: any, source: string) => void,
  notify: (msg: string, type: 'success' | 'error') => void, 
  updateLeadsConfig: (key: string, list: string[]) => Promise<void>,
  logLeadEvent: (leadId: string, type: string, description: string, metadata?: any) => Promise<void>,
  onSelectResearch?: (id: string) => void,
  users?: any[]
}) {
  const [viewMode, setViewMode] = useState<"board" | "table" | "insights" | "view">("board");
  const [ownerFilter, setOwnerFilter] = useState("All Owners");
  const isAdmin = userProfile?.role?.toLowerCase() === "admin";
  const isBoards = userProfile?.role?.toLowerCase() === "dcm boards" || isAdmin;
  const isAssessorRole = userProfile?.role?.toLowerCase() === "assessor";
  const isAssessor = isAssessorRole || isBoards;
  
  // Default permissions if not set
  const assessorPerms = settings?.assessorPermissions || {
    canAddLead: false,
    canEditLead: true,
    canDeleteLead: false,
    disallowedStatuses: ["Assessed", "Approved for future project"],
    canRunResearch: true
  };

  const isAdminOrBoard = isAdmin || isBoards;
  const canAddLead = isAdminOrBoard || (isAssessorRole && assessorPerms.canAddLead);
  const canEditLead = isAdminOrBoard || (isAssessorRole && assessorPerms.canEditLead);
  const canDeleteLead = isAdminOrBoard || (isAssessorRole && assessorPerms.canDeleteLead);
  const canRunResearch = isAdminOrBoard || (isAssessorRole && assessorPerms.canRunResearch);
  const canModifyStatus = isAdminOrBoard || isAssessor; // Base ability to change status, but restricted by options

  const [isEditing, setIsEditing] = useState(false);
  const [editingLead, setEditingLead] = useState<any | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const [csvText, setCsvText] = useState("");
  const [loading, setLoading] = useState(false);
  const [viewingLead, setViewingLead] = useState<any | null>(null);
  const [leadTimeline, setLeadTimeline] = useState<any[]>([]);

  const [approvalCommentLocal, setApprovalCommentLocal] = useState("");

  useEffect(() => {
    if (viewingLead) {
      setApprovalCommentLocal(viewingLead.approvalComment || "");
      
      // Fetch timeline
      const timelineRef = collection(db, "leads", viewingLead.id, "timeline");
      const q = query(timelineRef, orderBy("timestamp", "desc"));
      
      const unsubscribe = onSnapshot(q, (snapshot) => {
        const events = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        setLeadTimeline(events);
      }, (err) => handleFirestoreError(err, OperationType.LIST, `leads/${viewingLead.id}/timeline`));
      
      return () => unsubscribe();
    } else {
      setApprovalCommentLocal("");
      setLeadTimeline([]);
    }
  }, [viewingLead?.id]);

  const [newComment, setNewComment] = useState("");
  const [searchTerm, setSearchTerm] = useState("");
  const [isSuggestingEmail, setIsSuggestingEmail] = useState(false);
  const [suggestedEmail, setSuggestedEmail] = useState<string | null>(null);
  const [paidOnly, setPaidOnly] = useState(false);
  const [filterAssessment, setFilterAssessment] = useState(false);
  const [filterAssessed, setFilterAssessed] = useState(false);
  const [filterApproved, setFilterApproved] = useState(false);
  const [filterCompleted, setFilterCompleted] = useState(false);


  const leadStatuses = settings?.leadsConfig?.statuses || [
    "Under assessment",
    "Assessed",
    "Approved for future project",
    "Strong potential",
    "First meeting",
    "Needs identifying",
    "Under consideration",
    "Follow up",
    "Project completed"
  ];

  const leadSources = settings?.leadsConfig?.sources || [
    "LinkedIn",
    "Referral",
    "Direct Reachout",
    "Event",
    "Other"
  ];

  const leadTypes = settings?.leadsConfig?.types || ["Education", "Partnership", "Project"];

  const [selectedYear, setSelectedYear] = useState<string>("All Years");

  const availableYears = Array.from(new Set(leads.map(l => {
    const dateStr = l.lastContact || l.createdAt;
    if (!dateStr) return null;
    try {
      return new Date(dateStr).getFullYear().toString();
    } catch {
      return null;
    }
  }).filter(Boolean) as string[])).sort((a, b) => b.localeCompare(a));

  const [selectedStatuses, setSelectedStatuses] = useState<string[]>(() => {
    const saved = localStorage.getItem('dcm_selectedStatuses');
    return saved ? JSON.parse(saved) : [];
  });
  const [isStatusFilterOpen, setIsStatusFilterOpen] = useState(false);
  const [isEditingName, setIsEditingName] = useState(false);
  const [tempName, setTempName] = useState("");

  const [boardStatuses, setBoardStatuses] = useState<string[]>(leadStatuses);

  // List of statuses an Assessor is allowed to set
  const allowedStatusesForRole = isAssessorRole && !isAdminOrBoard
    ? leadStatuses.filter((s: string) => !(assessorPerms.disallowedStatuses || []).includes(s))
    : leadStatuses;

  // Sync board columns when settings change
  useEffect(() => {
    setBoardStatuses(leadStatuses);
  }, [settings?.leadsConfig?.statuses]);

  const [columnWidths, setColumnWidths] = useState<Record<string, number>>(() => {
    const globalWidths = settings?.leadsConfig?.columnWidths || {};
    const saved = localStorage.getItem('dcm_columnWidths');
    const localWidths = saved ? JSON.parse(saved) : {};
    return { ...localWidths, ...globalWidths };
  });

  const [activeLead, setActiveLead] = useState<any | null>(null);

  useEffect(() => {
    const globalWidths = settings?.leadsConfig?.columnWidths || {};
    if (Object.keys(globalWidths).length > 0) {
      setColumnWidths(prev => ({ ...prev, ...globalWidths }));
    }
  }, [settings?.leadsConfig?.columnWidths]);

  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [aiActions, setAiActions] = useState<LeadAction[]>([]);
  const [insightOwnerFilter, setInsightOwnerFilter] = useState<string>("All Owners");
  const [showAiInsights, setShowAiInsights] = useState(false); // Legacy modal state, keeping for safety but we will use viewMode

  useEffect(() => {
    localStorage.setItem('dcm_columnWidths', JSON.stringify(columnWidths));
  }, [columnWidths]);

  useEffect(() => {
    localStorage.setItem('dcm_boardStatuses', JSON.stringify(boardStatuses));
  }, [boardStatuses]);

  useEffect(() => {
    localStorage.setItem('dcm_selectedStatuses', JSON.stringify(selectedStatuses));
  }, [selectedStatuses]);

  const activeOwnerOptions = Array.from(new Set(
    leads
      .filter(l => {
        const s = (l.leadStatus || "").toLowerCase();
        return !["closed", "not confirmed", "not interested", "not suitable", "not interest"].includes(s);
      })
      .map(l => (l.leadOwner || "Unassigned").trim())
  ))
  .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }))
  .filter((v, i, a) => a.findIndex(t => t.toLowerCase() === v.toLowerCase()) === i);

  const assignableOwnerOptions = Array.from(new Set([
    ...leads.map(l => (l.leadOwner || "").trim()),
    ...users.map(u => (u.displayName || u.firstName || u.name || "").trim()).filter(Boolean)
  ]))
  .filter(Boolean)
  .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }))
  .filter((v, i, a) => a.findIndex(t => t.toLowerCase() === v.toLowerCase()) === i);

  // Load persisted insights from Firestore
  useEffect(() => {
    const q = query(collection(db, "leadInsights"), orderBy("createdAt", "desc"));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      if (!snapshot.empty) {
        // We just take the latest one
        const latestDoc = snapshot.docs[0].data();
        setAiActions(latestDoc.actions || []);
      }
    }, (err) => handleFirestoreError(err, OperationType.LIST, "leadInsights"));
    return () => unsubscribe();
  }, []);

  const handleRunAiAnalysis = async () => {
    setIsAnalyzing(true);
    try {
      const actions = await extractActionsFromLeads(leads, aiConfig);
      setAiActions(actions);
      
      // Persist to Firestore
      await addDoc(collection(db, "leadInsights"), {
        actions,
        createdAt: new Date().toISOString(),
        createdBy: auth.currentUser?.email || "System"
      });
      
      setViewMode("insights");
    } catch (err: any) {
      handleAiError(err, "Lead Analysis");
    } finally {
      setIsAnalyzing(false);
    }
  };

  const [columnDefs, setColumnDefs] = useState([
    { id: 'organisation', label: 'Organisation', width: 200 },
    { id: 'status', label: 'Status', width: 140 },
    { id: 'owner', label: 'Owner', width: 120 },
    { id: 'type', label: 'Type', width: 100 },
    { id: 'lastContact', label: 'Last Contact', width: 120 },
    { id: 'potentialDate', label: 'Potential Date', width: 120 },
    { id: 'actions', label: 'Actions', width: 100 }
  ]);
  const [isAdminOpen, setIsAdminOpen] = useState(false);

  const handleExportCSV = () => {
    if (leads.length === 0) return;
    
    // Prepare data for export
    const dataToExport = leads.map(l => ({
      Organisation: l.organisation || "",
      Status: l.leadStatus || "",
      ContactName: l.contactName || "",
      LeadOwner: l.leadOwner || "",
      Type: l.leadType || "",
      Email: l.email || "",
      Phone: l.phone || "",
      Location: l.location || "",
      Description: l.description || "",
      LinkedIn: l.linkedin || "",
      Website: l.website || "",
      LastContact: l.lastContactDate || "",
      PotentialDate: l.potentialProjectDate || "",
      ConfidenceScore: l.confidenceScore || "",
      StrategicAlignment: l.strategicAlignment || ""
    }));

    const csv = Papa.unparse(dataToExport);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `dcm_leads_export_${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: {
        distance: 8,
      },
    }),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  const handleDragStart = (event: DragStartEvent) => {
    if (event.active.data.current?.lead) {
      setActiveLead(event.active.data.current.lead);
    }
  };

  const handleDragEnd = async (event: DragEndEvent) => {
    const { active, over } = event;
    setActiveLead(null);

    if (!over) return;

    // Case Lead Dragged to Column
    if (active.data.current?.lead && over.data.current?.status) {
      const lead = active.data.current.lead;
      const newStatus = over.data.current.status;

      if (!canModifyStatus) {
        notify("You don't have permission to change status", "error");
        return;
      }

      // Check role-specific status restrictions
      if (isAssessorRole && !isAdminOrBoard && (assessorPerms.disallowedStatuses || []).includes(newStatus)) {
        notify(`Only DCM board can set leads to "${newStatus}"`, "error");
        return;
      }

      if (lead.leadStatus !== newStatus) {
        try {
          await setDoc(doc(db, "leads", lead.id), { 
            leadStatus: newStatus, 
            updatedAt: new Date().toISOString() 
          }, { merge: true });
          notify("Status updated", "success");
          if (logLeadEvent) {
            logLeadEvent(lead.id, "STATUS_CHANGE", `Lead moved to ${newStatus} via board`);
          }
        } catch (err) {
          notify("Update failed", "error");
          handleFirestoreError(err, OperationType.UPDATE, "leads");
        }
      }
      return;
    }

    // Case Column Reordering
    if (active.id !== over.id && boardStatuses.includes(active.id as string)) {
      const oldIndex = boardStatuses.indexOf(active.id as string);
      const newIndex = boardStatuses.indexOf(over.id as string);
      const newList = arrayMove(boardStatuses, oldIndex, newIndex);
      setBoardStatuses(newList);
      updateLeadsConfig("statuses", newList);
    }
  };

  const handleAddComment = async () => {
    if (!newComment.trim() || !viewingLead) return;
    const comment = {
      text: newComment,
      date: new Date().toLocaleDateString('en-US', { 
        month: 'long', 
        day: 'numeric', 
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      }),
      author: auth.currentUser?.displayName || "User"
    };
    
    const updatedLead = {
      ...viewingLead,
      notes: [...(viewingLead.notes || []), comment]
    };
    
    try {
      await setDoc(doc(db, "leads", viewingLead.id), updatedLead, { merge: true });
      await logLeadEvent(viewingLead.id, "NOTE_ADDED", "Internal note added", { text: newComment });
      setViewingLead(updatedLead);
      setNewComment("");
    } catch (err) {
      handleFirestoreError(err, OperationType.UPDATE, "leads");
    }
  };

  const handleSuggestFollowUp = async () => {
    if (!viewingLead) return;
    setIsSuggestingEmail(true);
    setSuggestedEmail(null);
    try {
      const draft = await suggestFollowUpEmail(viewingLead, userProfile, aiConfig);
      setSuggestedEmail(draft);
    } catch (err: any) {
      handleAiError(err, "Email Suggestions");
    } finally {
      setIsSuggestingEmail(false);
    }
  };

  const handleUpdateField = async (field: string, newValue: any) => {
    if (!viewingLead) return;
    
    // Role restriction
    if (field === "leadStatus") {
       if (!canModifyStatus) {
         notify("You don't have permission to change lead status.", "error");
         return;
       }
       if (isAssessorRole && !isAdminOrBoard && (assessorPerms.disallowedStatuses || []).includes(newValue)) {
         notify(`Only DCM board can set leads to "${newValue}"`, "error");
         return;
       }
    } else if (!canEditLead) {
      notify("You don't have permission to edit lead details.", "error");
      return;
    }

    const updatePayload = { [field]: newValue, updatedAt: new Date().toISOString() };
    try {
      await setDoc(doc(db, "leads", viewingLead.id), updatePayload, { merge: true });
      
      const updatedLead = { ...viewingLead, ...updatePayload };
      // Log event for major field changes
      if (field === "leadStatus") {
        await logLeadEvent(viewingLead.id, "STATUS_CHANGE", `Status changed to: ${newValue}`, { from: viewingLead.leadStatus, to: newValue });
      } else if (field === "leadOwner") {
        await logLeadEvent(viewingLead.id, "OWNER_CHANGE", `Owner changed to: ${newValue}`, { from: viewingLead.leadOwner, to: newValue });
      } else if (field === "isPaidEngagement") {
        await logLeadEvent(viewingLead.id, "ENGAGEMENT_CHANGE", `Engagement type changed to: ${newValue ? "Paid" : "Free"}`, { isPaid: newValue });
      }
      
      setViewingLead(updatedLead);
      notify("Field updated successfully", "success");
    } catch (err) {
      notify("Failed to update field (Permission Denied?)", "error");
      handleFirestoreError(err, OperationType.UPDATE, "leads");
    }
  };

  const handleTableUpdate = async (leadId: string, field: string, newValue: any) => {
    // Role restriction
    if (field === "leadStatus") {
       if (!canModifyStatus) {
         notify("You don't have permission to change lead status.", "error");
         return;
       }
       if (isAssessorRole && !isAdminOrBoard && (assessorPerms.disallowedStatuses || []).includes(newValue)) {
         notify(`Only DCM board can set leads to "${newValue}"`, "error");
         return;
       }
    }
    
    try {
      await setDoc(doc(db, "leads", leadId), { 
        [field]: newValue, 
        updatedAt: new Date().toISOString() 
      }, { merge: true });
      notify("Field updated", "success");
    } catch (err) {
      notify("Update failed", "error");
      handleFirestoreError(err, OperationType.UPDATE, "leads");
    }
  };

  const globalFilteredLeads = leads.filter(lead => {
    const dateStr = lead.lastContact || lead.createdAt;
    let yearMatches = true;
    if (selectedYear !== "All Years" && dateStr) {
      try {
        yearMatches = new Date(dateStr).getFullYear().toString() === selectedYear;
      } catch {
        yearMatches = false;
      }
    } else if (selectedYear !== "All Years") {
      yearMatches = false;
    }

    const matchesSearch = 
      lead.organisation?.toLowerCase().includes(searchTerm.toLowerCase()) ||
      lead.contactName?.toLowerCase().includes(searchTerm.toLowerCase()) ||
      lead.leadOwner?.toLowerCase().includes(searchTerm.toLowerCase());
    const currentLeadOwner = (lead.leadOwner || "Unassigned").trim().toLowerCase();
    const matchesOwner = ownerFilter === "All Owners" || currentLeadOwner === ownerFilter.toLowerCase();
    return matchesSearch && yearMatches && matchesOwner;
  });

  const filteredLeads = globalFilteredLeads.filter(lead => {
    const matchesPaid = !paidOnly || lead.isPaidEngagement;
    
    // Status Toggles logic
    const hasAnyStatusToggle = filterAssessment || filterAssessed || filterApproved || filterCompleted;
    let matchesStatusToggle = true;
    if (hasAnyStatusToggle) {
      matchesStatusToggle = false;
      const status = (lead.leadStatus || "").toLowerCase();
      if (filterAssessment && status.includes("under assessment")) matchesStatusToggle = true;
      if (filterAssessed && status === "assessed") matchesStatusToggle = true;
      if (filterApproved && status.includes("approved for future")) matchesStatusToggle = true;
      if (filterCompleted && (status === "completed" || status === "project completed")) matchesStatusToggle = true;
    }

    const matchesStatusSelection = selectedStatuses.length === 0 || selectedStatuses.includes(lead.leadStatus) || 
                                   selectedStatuses.some(s => s.toLowerCase() === (lead.leadStatus || "").toLowerCase());
    return matchesPaid && matchesStatusToggle && matchesStatusSelection;
  });

  // KPI Calculations based on global filters (Search + Year + Paid) but NOT specific status visibility
  const totalInDatabase = leads.length;
  const totalInView = globalFilteredLeads.length;
  const inAssessmentCount = globalFilteredLeads.filter(l => (l.leadStatus || "").toLowerCase().includes("under assessment")).length;
  const assessedCount = globalFilteredLeads.filter(l => (l.leadStatus || "").toLowerCase() === "assessed").length;
  const approvedCount = globalFilteredLeads.filter(l => (l.leadStatus || "").toLowerCase().includes("approved for future")).length;
  const paidCount = globalFilteredLeads.filter(l => l.isPaidEngagement).length;
  const completedCount = globalFilteredLeads.filter(l => (l.leadStatus || "").toLowerCase().includes("completed")).length;

  // Sort by Last Contact (recent at top)
  const sortedLeads = [...filteredLeads].sort((a, b) => {
    const dateA = a.lastContact ? new Date(a.lastContact).getTime() : 0;
    const dateB = b.lastContact ? new Date(b.lastContact).getTime() : 0;
    
    // Primary sort: Last Contact Date
    if (dateA !== dateB) return dateB - dateA;
    
    // Secondary sort: Created At if exists
    const createdA = a.createdAt ? new Date(a.createdAt).getTime() : 0;
    const createdB = b.createdAt ? new Date(b.createdAt).getTime() : 0;
    return createdB - createdA;
  });

  const handleLeadFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    setLoading(true);
    try {
      // In a real app, you'd upload to Storage. For now, we simulate by adding the filename to the attachments field.
      const currentAttachments = editingLead.attachments || "";
      const separator = currentAttachments ? ", " : "";
      setEditingLead({
        ...editingLead,
        attachments: currentAttachments + separator + file.name
      });
      // alert(`File "${file.name}" uploaded successfully (mock).`);
    } catch (err) {
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingLead.organisation?.trim()) {
      notify("Organisation name is required", "error");
      return;
    }
    
    setLoading(true);
    try {
      console.log("Saving lead. Role info:", { isAdminOrBoard, isAssessor, canModifyStatus });
      console.log("Original editingLead:", editingLead);
      
      let finalStatus = editingLead.leadStatus;
      if (!editingLead.id && isAssessor && !isAdminOrBoard) {
        console.log("Forcing status to 'Under assessment' for Assessor creation");
        finalStatus = "Under assessment";
      }

      // Duplicate check for new leads
      if (!editingLead.id) {
        const isDuplicate = leads.some(l => 
          l.organisation?.trim().toLowerCase() === editingLead.organisation?.trim().toLowerCase()
        );
        if (isDuplicate) {
          if (!confirm(`A lead with the name "${editingLead.organisation}" already exists. Do you want to create a duplicate?`)) {
            setLoading(false);
            return;
          }
        }
      }

      // Role restriction: check canModifyStatus
      if (editingLead.id && !canModifyStatus) {
        const originalLead = leads.find(l => l.id === editingLead.id);
        if (originalLead && originalLead.leadStatus !== editingLead.leadStatus) {
           notify("You don't have permission to change lead status.", "error");
           setLoading(false);
           return;
        }
      }

      // Sanitizing payload to remove undefined fields
      const payload = Object.fromEntries(
        Object.entries({ ...editingLead, leadStatus: finalStatus }).filter(([_, v]) => v !== undefined)
      );
      console.log("Final payload for Firestore:", payload);

      if (editingLead.id) {
        const id = editingLead.id;
        delete payload.id; // Don't save id field inside document
        await setDoc(doc(db, "leads", id), {
          ...payload,
          updatedAt: new Date().toISOString()
        }, { merge: true });
        await logLeadEvent(id, "UPDATE", "Lead details updated via form");
        notify("Lead updated successfully", "success");
      } else {
        const docRef = await addDoc(collection(db, "leads"), {
          ...payload,
          lastContact: payload.lastContact || new Date().toISOString().split('T')[0],
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        });
        await logLeadEvent(docRef.id, "CREATED", "Lead created in system");
        notify("Lead created successfully", "success");
      }
      
      setIsEditing(false);
      setEditingLead(null);
    } catch (err: any) {
      console.error("Critical Save Error:", err);
      const isPermissionError = err.message?.includes("permission-denied") || err.code === "permission-denied" || err.message?.includes("insufficient permissions");
      const errorMessage = isPermissionError 
        ? `Permission Denied: You cannot ${editingLead.id ? 'update' : 'create'} this lead. Ensure you have the correct role and fields.`
        : `Failed to save lead: ${err.message || 'Unknown error'}`;
      notify(errorMessage, "error");
      handleFirestoreError(err, editingLead.id ? OperationType.UPDATE : OperationType.CREATE, "leads");
    } finally {
      setLoading(false);
    }
  };

  const handleDelete = async (id: string) => {
    if (!canDeleteLead) {
      notify("You don't have permission to delete leads.", "error");
      return;
    }
    try {
      const leadToArchive = leads.find(l => l.id === id);
      if (leadToArchive) {
        await setDoc(doc(db, "archivedLeads", id), {
          ...leadToArchive,
          archivedAt: new Date().toISOString(),
          archivedBy: userProfile?.email
        });
      }
      
      const batch = writeBatch(db);
      batch.delete(doc(db, "leads", id));
      
      // Clean up orphaned assessments
      assessments.filter(a => a.nonProfitId === id).forEach(ass => {
        batch.delete(doc(db, "assessments", ass.id));
      });
      
      await batch.commit();
      notify("Lead moved to archive and related assessments removed", "success");
    } catch (err) {
      notify("Failed to archive lead (Permission Denied?)", "error");
      handleFirestoreError(err, OperationType.DELETE, "leads");
    }
  };

  const handleDuplicate = async (lead: any) => {
    if (!canAddLead) {
      notify("You don't have permission to duplicate leads.", "error");
      return;
    }
    setLoading(true);
    try {
      const { id, organisation, ...rest } = lead;
      const newLead = {
        ...rest,
        organisation: `${organisation} (Copy)`,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        createdBy: auth.currentUser?.email || "System",
        // Clear research data for the new context
        briefSummary: "",
        revenue: "",
        ein: "",
        charity_navigator_rating: "",
        propublica_grants: "",
        linkedin_overview: "",
        staff_linkedin_summary: "",
        staff_members: "",
        dcmComment: ""
      };
      const docRef = await addDoc(collection(db, "leads"), newLead);
      await logLeadEvent(docRef.id, "CREATED", `Lead duplicated from ${organisation}`);
      notify(`Duplicated ${organisation} for new project context.`, "success");
    } catch (err) {
      console.error(err);
      notify("Duplicate failed", "error");
    } finally {
      setLoading(false);
    }
  };

  const handleInvoiceUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 800000) {
        notify("File too large. Maximum size is 800KB for direct storage.", "error");
        return;
      }
      
      setLoading(true);
      const reader = new FileReader();
      reader.onload = async (event) => {
        const base64 = event.target?.result as string;
        setEditingLead({
          ...editingLead,
          invoiceName: file.name,
          invoiceData: base64,
          invoiceUploadedAt: new Date().toISOString()
        });
        setLoading(false);
        notify("Invoice ready. Click 'Save Changes' to finalize.", "success");
      };
      reader.onerror = () => {
        setLoading(false);
        notify("Failed to read file", "error");
      };
      reader.readAsDataURL(file);
    }
  };

  const handleDownloadTemplate = () => {
    const headers = [
      "Organisation", "Activity", "Attachments", "City", "Comments", 
      "Contact Name", "Country", "Direct link", "Donation", "Email", 
      "Last Contact", "Lead Type", "Lead owner", "Lead status", 
      "Project Languages", "Project Potential date", "Referred by", 
      "Website", "What's next?"
    ];
    const csvContent = headers.join(",") + "\n";
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = "dcm_leads_template.csv";
    a.click();
    window.URL.revokeObjectURL(url);
  };

  const handleImport = async () => {
    if (!csvText) return;
    setLoading(true);
    try {
      Papa.parse(csvText, {
        header: true,
        skipEmptyLines: true,
        transformHeader: (header) => header.trim(),
        complete: async (results) => {
          const rows = results.data;
          const CHUNK_SIZE = 400;
          
          try {
            for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
              const chunk = rows.slice(i, i + CHUNK_SIZE);
              const batch = writeBatch(db);
              
              chunk.forEach((row: any) => {
                const orgName = (String(row["Organisation"] || "")).trim().substring(0, 499);
                if (!orgName) return; 

                const rawStatus = (row["Lead status"] || "").trim().toLowerCase();
                const matchedStatus = boardStatuses.find(s => s.toLowerCase().trim() === rawStatus) || 
                                     (rawStatus.includes("closed") ? "Closed" : 
                                      rawStatus.includes("completed") ? "Project Completed" :
                                      rawStatus.includes("not interest") ? "Not Interested" :
                                      "Under assessment");

                const docRef = doc(collection(db, "leads"));
                const today = new Date().toISOString().split('T')[0];
                batch.set(docRef, {
                  organisation: orgName,
                  activity: row["Activity"] || "",
                  attachments: row["Attachments"] || "",
                  city: row["City"] || "",
                  comments: row["Comments"] || "",
                  contactName: row["Contact Name"] || "",
                  country: row["Country"] || "",
                  directLink: row["Direct link"] || "",
                  donation: row["Donation"] || "",
                  email: row["Email"] || "",
                  lastContact: row["Last Contact"] || today,
                  leadType: row["Lead Type"] || "",
                  leadOwner: row["Lead owner"] || "",
                  leadStatus: matchedStatus,
                  projectLanguages: row["Project Languages"] || "",
                  projectPotentialDate: row["Project Potential date"] || "",
                  referredBy: row["Referred by"] || "",
                  website: row["Website"] || "",
                  whatsNext: row["What's next?"] || "",
                  createdAt: new Date().toISOString(),
                  updatedAt: new Date().toISOString()
                });
              });
              
              await batch.commit();
            }
            setIsImporting(false);
            setCsvText("");
          } catch (err) {
            console.error("Batch commit error:", err);
          } finally {
            setLoading(false);
          }
        }
      });
    } catch (err) {
      console.error("Import error:", err);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="space-y-6">
      {/* Top Right Filters */}
      <div className="flex justify-end items-center gap-6 mb-2">
        {/* Year Filter */}
        <div className="flex items-center gap-3 bg-white border border-brand-border rounded-xl px-4 py-1.5 shadow-sm">
          <Calendar className="w-3.5 h-3.5 text-brand-muted" />
          <select 
            value={selectedYear}
            onChange={(e) => setSelectedYear(e.target.value)}
            className="bg-transparent text-[10px] font-black uppercase tracking-widest outline-none cursor-pointer h-7 min-w-[90px]"
          >
            <option value="All Years">All Years</option>
            {availableYears.map(year => (
              <option key={year} value={year}>{year}</option>
            ))}
          </select>
        </div>

        <div className="relative group">
          <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-brand-muted group-focus-within:text-brand-accent transition-colors" />
          <input 
            type="text" 
            placeholder="Search leads..." 
            className="pl-10 pr-6 py-2.5 bg-white border border-brand-border rounded-xl text-xs font-semibold focus:ring-2 focus:ring-brand-accent/20 outline-none w-72 shadow-sm transition-all"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
      </div>

      {/* KPI Section */}
      <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-6 gap-6 mb-8">
        <StatsCard icon={Users} label="Total CRM" value={totalInDatabase.toString()} subLabel="Full Database" />
        <StatsCard 
          icon={Activity} 
          label="In Assessment" 
          value={inAssessmentCount.toString()} 
          subLabel="Active Pipeline" 
          onClick={() => setFilterAssessment(!filterAssessment)}
          isActive={filterAssessment}
        />
        <StatsCard 
          icon={CheckCircle2} 
          label="Assessed" 
          value={assessedCount.toString()} 
          subLabel="Ready Review" 
          onClick={() => setFilterAssessed(!filterAssessed)}
          isActive={filterAssessed}
        />
        <StatsCard 
          icon={Target} 
          label="Future Projects" 
          value={approvedCount.toString()} 
          subLabel="Approved" 
          onClick={() => setFilterApproved(!filterApproved)}
          isActive={filterApproved}
        />
        <StatsCard 
          icon={DollarSign} 
          label="Paid Leads" 
          value={paidCount.toString()} 
          subLabel="Engagements" 
          onClick={() => setPaidOnly(!paidOnly)}
          isActive={paidOnly}
        />
        <StatsCard 
          icon={CalendarDays} 
          label="Completed" 
          value={completedCount.toString()} 
          subLabel="Verified" 
          onClick={() => setFilterCompleted(!filterCompleted)}
          isActive={filterCompleted}
        />
      </div>

      <div className="flex items-center justify-between mb-8">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 bg-white rounded-lg p-1 border border-brand-border">
            <button 
              onClick={() => setViewMode("board")}
              className={cn(
                "px-4 py-2 rounded-md text-xs font-bold uppercase tracking-widest transition-all",
                viewMode === "board" ? "bg-brand-primary text-white" : "text-brand-muted hover:bg-brand-bg"
              )}
            >
              Board
            </button>
            <button 
              onClick={() => setViewMode("table")}
              className={cn(
                "px-4 py-2 rounded-md text-xs font-bold uppercase tracking-widest transition-all",
                viewMode === "table" ? "bg-brand-primary text-white" : "text-brand-muted hover:bg-brand-bg"
              )}
            >
              Table
            </button>
            <button 
              onClick={() => setViewMode("insights")}
              className={cn(
                "px-4 py-2 rounded-md text-xs font-bold uppercase tracking-widest transition-all",
                viewMode === "insights" ? "bg-brand-primary text-white" : "text-brand-muted hover:bg-brand-bg"
              )}
            >
              Insights
            </button>
          </div>

          <div className="h-6 w-px bg-brand-border mx-2" />
        </div>
        <div className="flex items-center gap-3">
          {viewMode !== "insights" && (
            <>
              {/* Lead Owner Filter */}
              <div className="flex items-center gap-3 bg-white border border-brand-border rounded-xl px-4 py-1.5 shadow-sm">
                <Users className="w-3.5 h-3.5 text-brand-muted" />
                <select 
                  value={ownerFilter}
                  onChange={(e) => setOwnerFilter(e.target.value)}
                  className="bg-transparent text-[10px] font-black uppercase tracking-widest outline-none cursor-pointer h-7 min-w-[120px]"
                >
                  <option value="All Owners">All Owners</option>
                  {activeOwnerOptions.map(owner => (
                    <option key={owner} value={owner}>{owner}</option>
                  ))}
                </select>
              </div>

              {paidOnly && (
                <button 
                  onClick={() => setPaidOnly(false)}
                  className="px-4 py-2.5 rounded-xl bg-brand-accent/10 text-brand-accent border border-brand-accent/20 transition-all flex items-center gap-2 text-[10px] font-black uppercase tracking-widest shadow-sm"
                >
                  <X className="w-3.5 h-3.5" /> Paid Only
                </button>
              )}

              {filterAssessment && (
                <button 
                  onClick={() => setFilterAssessment(false)}
                  className="px-4 py-2.5 rounded-xl bg-brand-bg text-brand-primary border border-brand-border transition-all flex items-center gap-2 text-[10px] font-black uppercase tracking-widest shadow-sm"
                >
                  <X className="w-3.5 h-3.5" /> In Assessment
                </button>
              )}

              {filterAssessed && (
                <button 
                  onClick={() => setFilterAssessed(false)}
                  className="px-4 py-2.5 rounded-xl bg-brand-bg text-brand-primary border border-brand-border transition-all flex items-center gap-2 text-[10px] font-black uppercase tracking-widest shadow-sm"
                >
                  <X className="w-3.5 h-3.5" /> Assessed
                </button>
              )}

              {filterApproved && (
                <button 
                  onClick={() => setFilterApproved(false)}
                  className="px-4 py-2.5 rounded-xl bg-brand-bg text-brand-primary border border-brand-border transition-all flex items-center gap-2 text-[10px] font-black uppercase tracking-widest shadow-sm"
                >
                  <X className="w-3.5 h-3.5" /> Approved
                </button>
              )}

              {filterCompleted && (
                <button 
                  onClick={() => setFilterCompleted(false)}
                  className="px-4 py-2.5 rounded-xl bg-brand-bg text-brand-primary border border-brand-border transition-all flex items-center gap-2 text-[10px] font-black uppercase tracking-widest shadow-sm"
                >
                  <X className="w-3.5 h-3.5" /> Completed
                </button>
              )}


              
              <div className="relative">
                <button 
                  onClick={() => setIsStatusFilterOpen(!isStatusFilterOpen)}
                  className="bg-white border text-[10px] border-brand-border rounded-xl px-4 py-2.5 font-black uppercase tracking-widest outline-none cursor-pointer focus:ring-2 focus:ring-brand-accent/20 shadow-sm flex items-center gap-2 min-w-[160px] justify-between"
                >
                  <span className="truncate">
                    {selectedStatuses.length === 0 ? "All Statuses" : 
                     selectedStatuses.length === 1 ? selectedStatuses[0] : 
                     `${selectedStatuses.length} Statuses`}
                  </span>
                  <ChevronRight className={cn("w-3 h-3 transition-transform", isStatusFilterOpen ? "rotate-90" : "")} />
                </button>
                
                <AnimatePresence>
                  {isStatusFilterOpen && (
                    <>
                      <div className="fixed inset-0 z-[70]" onClick={() => setIsStatusFilterOpen(false)} />
                      <motion.div 
                        initial={{ opacity: 0, y: 10, scale: 0.95 }}
                        animate={{ opacity: 1, y: 0, scale: 1 }}
                        exit={{ opacity: 0, y: 10, scale: 0.95 }}
                        className="absolute right-0 top-full mt-2 w-72 bg-white border border-brand-border rounded-2xl shadow-2xl z-[80] overflow-hidden"
                      >
                        <div className="p-3 border-b border-brand-bg flex items-center justify-between bg-brand-bg/30">
                          <span className="text-[10px] font-black uppercase tracking-widest text-brand-muted">Filter Status</span>
                          <div className="flex gap-2">
                            <button 
                              onClick={() => setSelectedStatuses([])}
                              className="text-[9px] font-bold text-brand-accent hover:underline uppercase tracking-widest"
                            >
                              Clear
                            </button>
                            <button 
                              onClick={() => setSelectedStatuses([...boardStatuses])}
                              className="text-[9px] font-bold text-brand-accent hover:underline uppercase tracking-widest"
                            >
                              All
                            </button>
                          </div>
                        </div>
                        <div className="max-h-64 overflow-y-auto py-2">
                          {boardStatuses.map(s => (
                            <label 
                              key={s} 
                              className="flex items-center gap-3 px-4 py-2.5 hover:bg-brand-bg cursor-pointer group transition-colors"
                            >
                              <div className={cn(
                                "w-4 h-4 rounded border flex items-center justify-center transition-all",
                                selectedStatuses.includes(s) ? "bg-brand-accent border-brand-accent" : "border-brand-border group-hover:border-brand-accent"
                              )}>
                                {selectedStatuses.includes(s) && <CheckCircle2 className="w-3 h-3 text-white" />}
                              </div>
                              <input 
                                type="checkbox"
                                className="hidden"
                                checked={selectedStatuses.includes(s)}
                                onChange={() => {
                                  if (selectedStatuses.includes(s)) {
                                    setSelectedStatuses(selectedStatuses.filter(item => item !== s));
                                  } else {
                                    setSelectedStatuses([...selectedStatuses, s]);
                                  }
                                }}
                              />
                              <span className={cn(
                                "text-[11px] font-black uppercase tracking-widest",
                                selectedStatuses.includes(s) ? "text-brand-primary" : "text-brand-muted"
                              )}>
                                {s}
                              </span>
                            </label>
                          ))}
                        </div>
                      </motion.div>
                    </>
                  )}
                </AnimatePresence>
              </div>
            </>
          )}
          <button 
            onClick={handleRunAiAnalysis}
            disabled={isAnalyzing || !canRunResearch}
            className="px-6 py-2.5 bg-gradient-to-r from-brand-accent to-blue-600 text-white rounded-lg font-bold text-xs uppercase tracking-widest hover:opacity-90 transition-all flex items-center gap-2 shadow-lg shadow-brand-accent/20 disabled:opacity-50"
          >
            {isAnalyzing ? <Loader2 className="w-4 h-4 animate-spin" /> : <ShieldCheck className="w-4 h-4" />}
            {aiActions.length > 0 ? "Refresh Insights" : "AI Insights"}
          </button>
          {viewMode !== "insights" && canAddLead && (
              <button 
                onClick={() => {
                  setEditingLead({ 
                    organisation: "", 
                    leadStatus: "Under assessment",
                    lastContact: new Date().toISOString().split('T')[0]
                  });
                  setIsEditing(true);
                }}
                className="px-6 py-2.5 bg-brand-accent text-white rounded-lg font-bold text-xs uppercase tracking-widest hover:opacity-90 transition-all flex items-center gap-2"
              >
                <Plus className="w-4 h-4" /> Add Lead
              </button>
          )}
        </div>
      </div>

      {isAdminOpen && (
        <div className="fixed inset-0 bg-brand-primary/60 backdrop-blur-sm z-[60] flex items-center justify-center p-6">
          <motion.div 
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            className="bg-white rounded-3xl w-full max-w-lg shadow-2xl overflow-hidden border border-brand-border"
          >
            <div className="p-8 border-b border-brand-border bg-brand-bg flex justify-between items-center">
              <div>
                <h3 className="text-xl font-black text-brand-primary uppercase tracking-tight">{isAdmin ? "Admin Area" : "Settings"}</h3>
                <p className="text-[10px] font-bold text-brand-muted uppercase tracking-widest mt-1">{isAdmin ? "Manage pipeline structure and data" : "Customize your pipeline view"}</p>
              </div>
              <button 
                onClick={() => setIsAdminOpen(false)}
                className="p-2 hover:bg-brand-border rounded-full transition-colors text-brand-muted"
              >
                <X className="w-6 h-6" />
              </button>
            </div>
            <div className="p-8 max-h-[70vh] overflow-y-auto custom-scrollbar space-y-10">
              <section className="space-y-6">
                <div>
                  <div className="flex items-center justify-between mb-4">
                    <div>
                      <h4 className="text-xs font-black text-brand-primary uppercase tracking-widest">Pipeline Stages</h4>
                      <p className="text-[9px] font-bold text-brand-muted uppercase mt-1">Reorder, rename, or resize columns</p>
                    </div>
                    {isAdmin && (
                      <button 
                        onClick={() => {
                          const newList = [...boardStatuses, "New Status"];
                          setBoardStatuses(newList);
                          updateLeadsConfig("statuses", newList);
                        }}
                        className="p-2 border border-brand-accent/20 bg-brand-accent/5 text-brand-accent rounded-lg hover:bg-brand-accent/10 transition-colors"
                      >
                        <Plus className="w-4 h-4" />
                      </button>
                    )}
                  </div>
                  <DndContext 
                    sensors={sensors}
                    collisionDetection={closestCenter}
                    onDragEnd={handleDragEnd}
                  >
                    <SortableContext 
                      items={boardStatuses}
                      strategy={verticalListSortingStrategy}
                    >
                      <div className="space-y-3">
                        {boardStatuses.map(status => (
                          <SortableStatusItem 
                            key={status} 
                            id={status} 
                            boardStatuses={boardStatuses}
                            setBoardStatuses={setBoardStatuses}
                            columnWidths={columnWidths}
                            setColumnWidths={setColumnWidths}
                            onUpdateGlobal={(newList) => updateLeadsConfig("statuses", newList)}
                            readOnly={!isAdmin}
                          />
                        ))}
                      </div>
                    </SortableContext>
                  </DndContext>
                </div>
              </section>
            </div>
            <div className="p-8 border-t border-brand-border bg-brand-bg flex justify-end">
              <button 
                onClick={() => setIsAdminOpen(false)}
                className="px-8 py-3 bg-brand-primary text-white rounded-xl font-black uppercase tracking-widest text-[10px] shadow-lg shadow-brand-primary/20 hover:opacity-90 transition-all"
              >
                Close & Save Settings
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {isImporting && (
        <div className="fixed inset-0 bg-brand-primary/60 backdrop-blur-sm z-50 flex items-center justify-center p-6">
          <div className="bg-white rounded-2xl w-full max-w-2xl shadow-2xl overflow-hidden border border-brand-border">
            <div className="p-8 border-b border-brand-border bg-brand-bg flex justify-between items-center">
              <div>
                <h3 className="text-xl font-black text-brand-primary uppercase tracking-tight">Import DCM Leads</h3>
                <p className="text-xs font-bold text-brand-muted uppercase tracking-widest mt-1">Paste CSV records from your file</p>
              </div>
              <button 
                onClick={() => setIsImporting(false)}
                className="p-2 hover:bg-brand-border rounded-full transition-colors text-brand-muted"
              >
                <X className="w-6 h-6" />
              </button>
            </div>
            <div className="p-8 space-y-6">
              <div className="flex justify-between items-center mb-4">
                <p className="text-xs font-bold text-brand-muted uppercase tracking-widest">Paste CSV data below</p>
                <button 
                  onClick={handleDownloadTemplate}
                  className="flex items-center gap-2 text-[10px] font-black uppercase tracking-widest text-brand-accent hover:text-brand-accent/80 transition-colors"
                >
                  <Download className="w-3.5 h-3.5" /> Download Template
                </button>
              </div>
              <textarea 
                className="w-full h-80 bg-brand-bg border border-brand-border rounded-xl p-6 font-mono text-xs focus:ring-2 focus:ring-brand-accent/20 outline-none resize-none"
                placeholder="Paste CSV here (Organisation, Activity, ...)"
                value={csvText}
                onChange={(e) => setCsvText(e.target.value)}
              />
              <div className="flex justify-end gap-3 pt-4">
                <button 
                  onClick={() => setIsImporting(false)}
                  className="px-8 py-3 font-bold text-brand-muted uppercase tracking-widest text-xs hover:text-brand-primary"
                >
                  Cancel
                </button>
                <button 
                  onClick={handleImport}
                  disabled={loading || !csvText}
                  className="px-8 py-3 bg-brand-accent text-white rounded-lg font-bold uppercase tracking-widest text-xs shadow-lg shadow-brand-accent/20 hover:opacity-90 transition-all disabled:opacity-50 flex items-center gap-2"
                >
                  {loading && <Loader2 className="w-4 h-4 animate-spin" />}
                  Start Import
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {isEditing && (
        <div className="fixed inset-0 bg-brand-primary/60 backdrop-blur-sm z-50 flex items-center justify-center p-6 bg-scroll">
          <div className="bg-white rounded-2xl w-full max-w-3xl shadow-2xl overflow-hidden border border-brand-border max-h-[90vh] overflow-y-auto">
            <div className="p-8 border-b border-brand-border bg-brand-bg flex justify-between items-center sticky top-0 z-10">
              <div>
                <div className="flex flex-col">
                  <h3 className="text-xl font-black text-brand-primary uppercase tracking-tight">
                    {editingLead.id ? "Edit Lead" : "New Lead"}
                  </h3>
                  {editingLead.needsFurtherAssessment && (
                    <div className="flex items-center gap-2 mt-1 bg-amber-50 px-2 py-1 rounded w-fit border border-amber-100">
                      <AlertCircle className="w-3 h-3 text-amber-500" />
                      <p className="text-[9px] font-bold text-amber-600 uppercase tracking-widest">
                        Needs Further Assessment
                      </p>
                      <button 
                        type="button"
                        onClick={() => setEditingLead({...editingLead, needsFurtherAssessment: false})}
                        className="ml-2 text-[8px] font-black underline text-amber-700 hover:text-amber-800"
                      >
                        Clear Flag
                      </button>
                    </div>
                  )}
                  {editingLead.assessedBy && !editingLead.needsFurtherAssessment && (
                    <div className="flex items-center gap-2 mt-1 bg-emerald-50 px-2 py-1 rounded w-fit border border-emerald-100">
                      <CheckCircle2 className="w-3 h-3 text-emerald-500" />
                      <p className="text-[9px] font-bold text-emerald-600 uppercase tracking-widest">
                        Assessed by {(editingLead.assessedBy || "").split('@')[0]} on {editingLead.assessedAt ? new Date(editingLead.assessedAt).toLocaleDateString() : "Pending"}
                      </p>
                    </div>
                  )}
                </div>
              </div>
              <button onClick={() => setIsEditing(false)} className="p-2 hover:bg-brand-border rounded-full transition-colors text-brand-muted">
                <X className="w-6 h-6" />
              </button>
            </div>
            <form onSubmit={handleSave} className="p-8 grid grid-cols-2 gap-6 pb-20">
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Organisation</label>
                <input 
                  required
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                  value={editingLead.organisation || ''}
                  onChange={(e) => setEditingLead({...editingLead, organisation: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Lead Status</label>
                <select 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none appearance-none cursor-pointer"
                  value={editingLead.leadStatus || ''}
                  onChange={(e) => setEditingLead({...editingLead, leadStatus: e.target.value})}
                >
                  {allowedStatusesForRole.map((s: string) => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Lead Type</label>
                <select 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none appearance-none cursor-pointer"
                  value={editingLead.leadType || ''}
                  onChange={(e) => setEditingLead({...editingLead, leadType: e.target.value})}
                >
                  <option value="">Select Type...</option>
                  {leadTypes.map((t: string) => <option key={t} value={t}>{t}</option>)}
                </select>
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Lead Source</label>
                <select 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none appearance-none cursor-pointer"
                  value={editingLead.leadSource || ''}
                  onChange={(e) => setEditingLead({...editingLead, leadSource: e.target.value})}
                >
                  <option value="">Select Source...</option>
                  {leadSources.map((s: string) => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Lead Owner</label>
                <input 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                  value={editingLead.leadOwner || ''}
                  onChange={(e) => setEditingLead({...editingLead, leadOwner: e.target.value})}
                />
              </div>
              <div className="col-span-2 space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Activity / Description</label>
                <textarea 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none h-24"
                  value={editingLead.activity || ''}
                  onChange={(e) => setEditingLead({...editingLead, activity: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Contact Name</label>
                <input 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                  value={editingLead.contactName || ''}
                  onChange={(e) => setEditingLead({...editingLead, contactName: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Email</label>
                <input 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                  value={editingLead.email || ''}
                  onChange={(e) => setEditingLead({...editingLead, email: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Country</label>
                <input 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                  value={editingLead.country || ''}
                  onChange={(e) => setEditingLead({...editingLead, country: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">City</label>
                <input 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                  value={editingLead.city || ''}
                  onChange={(e) => setEditingLead({...editingLead, city: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Website</label>
                <input 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                  value={editingLead.website || ''}
                  onChange={(e) => setEditingLead({...editingLead, website: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Direct Link</label>
                <input 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                  value={editingLead.directLink || ''}
                  onChange={(e) => setEditingLead({...editingLead, directLink: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Referred by</label>
                <input 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                  value={editingLead.referredBy || ''}
                  onChange={(e) => setEditingLead({...editingLead, referredBy: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Last Contact Date</label>
                <input 
                  type="date"
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                  value={editingLead.lastContact || ''}
                  onChange={(e) => setEditingLead({...editingLead, lastContact: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Project Potential Date</label>
                <input 
                  type="date"
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                  value={editingLead.projectPotentialDate || ''}
                  onChange={(e) => setEditingLead({...editingLead, projectPotentialDate: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Project Languages</label>
                <input 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                  value={editingLead.projectLanguages || ''}
                  onChange={(e) => setEditingLead({...editingLead, projectLanguages: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">What's Next?</label>
                <input 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                  value={editingLead.whatsNext || ''}
                  onChange={(e) => setEditingLead({...editingLead, whatsNext: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest text-emerald-600">Assessed By</label>
                <input 
                  className="w-full bg-emerald-50 border border-emerald-100 rounded-lg px-4 py-2.5 font-bold text-emerald-700 focus:ring-2 focus:ring-emerald-500/20 outline-none"
                  placeholder="Not assessed yet"
                  value={editingLead.assessedBy || ''}
                  onChange={(e) => setEditingLead({...editingLead, assessedBy: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest text-emerald-600">Assessed Date</label>
                <input 
                  type="date"
                  className="w-full bg-emerald-50 border border-emerald-100 rounded-lg px-4 py-2.5 font-bold text-emerald-700 focus:ring-2 focus:ring-emerald-500/20 outline-none"
                  value={editingLead.assessedAt ? new Date(editingLead.assessedAt).toISOString().split('T')[0] : ''}
                  onChange={(e) => setEditingLead({...editingLead, assessedAt: e.target.value ? new Date(e.target.value).toISOString() : null})}
                />
              </div>

              <div className="col-span-2 border-t border-brand-border pt-6 mt-4">
                <h4 className="text-[10px] font-black text-brand-primary uppercase tracking-[0.2em] mb-4 flex items-center gap-2">
                  <Database className="w-3.5 h-3.5 text-blue-500" /> Research Intelligence
                </h4>
                <div className="grid grid-cols-2 gap-6">
                  <div className="space-y-2">
                    <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Financial Tier / Revenue</label>
                    <input 
                      className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-blue-600 outline-none"
                      value={editingLead.revenue || ''}
                      onChange={(e) => setEditingLead({...editingLead, revenue: e.target.value})}
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">EIN Number</label>
                    <input 
                      className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-blue-600 outline-none"
                      value={editingLead.ein || ''}
                      onChange={(e) => setEditingLead({...editingLead, ein: e.target.value})}
                    />
                  </div>
                  <div className="col-span-2 space-y-2">
                    <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">DCM Brief Summary</label>
                    <textarea 
                      className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-blue-600 outline-none h-24"
                      value={editingLead.briefSummary || ''}
                      onChange={(e) => setEditingLead({...editingLead, briefSummary: e.target.value})}
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Charity Navigator Rating</label>
                    <input 
                      className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-medium text-brand-muted outline-none"
                      value={editingLead.charity_navigator_rating || ''}
                      onChange={(e) => setEditingLead({...editingLead, charity_navigator_rating: e.target.value})}
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">LinkedIn Overview</label>
                    <textarea 
                      className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-medium text-brand-muted outline-none h-20"
                      value={editingLead.linkedin_overview || ''}
                      onChange={(e) => setEditingLead({...editingLead, linkedin_overview: e.target.value})}
                    />
                  </div>
                </div>
              </div>

              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Donation</label>
                <input 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                  value={editingLead.donation || ''}
                  onChange={(e) => setEditingLead({...editingLead, donation: e.target.value})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Attachments</label>
                <div className="flex gap-2">
                  <input 
                    className="flex-1 bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                    value={editingLead.attachments || ''}
                    onChange={(e) => setEditingLead({...editingLead, attachments: e.target.value})}
                    placeholder="Filenames or links"
                  />
                  <label className="px-4 py-2.5 bg-brand-border text-brand-primary rounded-lg font-bold text-[10px] uppercase tracking-widest hover:bg-brand-border/80 transition-all cursor-pointer flex items-center gap-2">
                    <UploadCloud className="w-4 h-4" />
                    Upload
                    <input type="file" className="hidden" onChange={handleLeadFileUpload} />
                  </label>
                </div>
              </div>
              <div className="col-span-2 space-y-2">
                <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">General Comments</label>
                <textarea 
                  className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none h-20"
                  value={editingLead.comments || ''}
                  onChange={(e) => setEditingLead({...editingLead, comments: e.target.value})}
                />
              </div>

              {/* Paid Engagement Section */}
              <div className="col-span-2 border-t border-brand-border pt-6 mt-2 space-y-6">
                <div>
                  <h4 className="text-[10px] font-black text-brand-primary uppercase tracking-[0.2em] mb-4">Financials & Engagement</h4>
                  <div className="grid grid-cols-2 gap-6">
                    <div className="flex items-center gap-3 bg-brand-bg p-4 rounded-xl border border-brand-border">
                      <input 
                        type="checkbox"
                        id="paidEng"
                        className="w-4 h-4 rounded border-brand-border text-brand-accent focus:ring-brand-accent"
                        checked={editingLead.isPaidEngagement || false}
                        onChange={(e) => setEditingLead({...editingLead, isPaidEngagement: e.target.checked})}
                      />
                      <label htmlFor="paidEng" className="text-xs font-bold text-brand-primary cursor-pointer select-none">
                        Paid Engagement
                      </label>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Fees Paid ($CAD)</label>
                      <input 
                        className="w-full bg-brand-bg border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary focus:ring-2 focus:ring-brand-accent/20 outline-none"
                        value={editingLead.feesPaid || ''}
                        onChange={(e) => setEditingLead({...editingLead, feesPaid: e.target.value})}
                        placeholder="e.g. 5,000"
                      />
                    </div>
                    <div className="col-span-2 space-y-2">
                      <label className="text-[10px] font-bold text-brand-muted uppercase tracking-widest">Invoice Management</label>
                      <div className="flex gap-2">
                        <input 
                          disabled
                          className="flex-1 bg-brand-bg/50 border border-brand-border rounded-lg px-4 py-2.5 font-bold text-brand-primary/50 text-[10px] truncate"
                          value={editingLead.invoiceName ? `File: ${editingLead.invoiceName} (Uploaded ${new Date(editingLead.invoiceUploadedAt).toLocaleDateString()})` : "No invoice uploaded."}
                        />
                        <label className="px-4 py-2.5 bg-brand-accent text-white rounded-lg font-bold text-[10px] uppercase tracking-widest hover:opacity-90 transition-all cursor-pointer flex items-center gap-2">
                          <UploadCloud className="w-4 h-4" />
                          {editingLead.invoiceName ? "Update Invoice" : "Upload Invoice"}
                          <input 
                            type="file" 
                            className="hidden" 
                            onChange={handleInvoiceUpload} 
                          />
                        </label>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              <div className="col-span-2 flex justify-end gap-3 pt-6 border-t border-brand-border mt-4">
                <button 
                  type="button" 
                  onClick={() => setIsEditing(false)}
                  className="px-8 py-3 font-bold text-brand-muted uppercase tracking-widest text-xs hover:text-brand-primary"
                >
                  Cancel
                </button>
                <button 
                  type="submit"
                  disabled={loading}
                  className="px-8 py-3 bg-brand-accent text-white rounded-lg font-bold uppercase tracking-widest text-xs shadow-lg hover:opacity-90 transition-all flex items-center gap-2"
                >
                  {loading && <Loader2 className="w-4 h-4 animate-spin" />}
                  Save Changes
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      <AnimatePresence mode="wait">
        {viewMode === "insights" ? (
          <motion.div 
            key="insights"
            initial={{ opacity: 0, x: 20 }}
            animate={{ opacity: 1, x: 0 }}
            exit={{ opacity: 0, x: -20 }}
            className="bg-[#1A1A1A] rounded-3xl w-full shadow-2xl overflow-hidden border border-white/5 flex flex-col min-h-[600px]"
          >
            <div className="p-8 border-b border-white/5 bg-white/5 flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
              <div className="flex items-center gap-4">
                <div className="w-12 h-12 rounded-2xl bg-brand-accent/20 flex items-center justify-center">
                  <ShieldCheck className="w-6 h-6 text-brand-accent" />
                </div>
                <div>
                  <h3 className="text-xl font-black text-white uppercase tracking-tight">AI Insights & Reminders</h3>
                  <p className="text-[10px] font-bold text-white/40 uppercase tracking-widest mt-1">Extracted from comments and internal notes</p>
                </div>
              </div>
              
              <div className="flex flex-wrap items-center gap-4 w-full md:w-auto">
                <div className="relative flex-1 md:flex-initial min-w-[200px]">
                  <div className="absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none">
                    <UserCircle className="w-4 h-4 text-white/40" />
                  </div>
                  <select
                    className="w-full bg-white/5 border border-white/10 rounded-xl pl-10 pr-10 py-2.5 text-[10px] font-bold text-white uppercase tracking-widest focus:ring-1 focus:ring-brand-accent transition-all outline-none appearance-none"
                    value={insightOwnerFilter}
                    onChange={(e) => setInsightOwnerFilter(e.target.value)}
                  >
                    <option value="All Owners" className="bg-[#1A1A1A]">All Owners</option>
                    {Array.from(new Set(aiActions.map(a => a.leadOwner))).filter(Boolean).map(owner => (
                      <option key={owner} value={owner} className="bg-[#1A1A1A]">{owner}</option>
                    ))}
                  </select>
                  <div className="absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none">
                    <ChevronDown className="w-3 h-3 text-white/40" />
                  </div>
                </div>

                {isAnalyzing ? (
                  <div className="flex items-center gap-2 text-brand-accent animate-pulse bg-brand-accent/10 px-4 py-2.5 rounded-xl border border-brand-accent/20">
                    <Loader2 className="w-4 h-4 animate-spin" />
                    <span className="text-[10px] font-bold uppercase tracking-widest">Updating...</span>
                  </div>
                ) : (
                  <button 
                    onClick={handleRunAiAnalysis}
                    className="bg-brand-accent text-white px-4 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest hover:scale-[1.02] active:scale-[0.98] transition-all shadow-lg flex items-center gap-2"
                  >
                    Refresh Analysis
                  </button>
                )}
              </div>
            </div>
            
            <div className="p-8 space-y-4">
              {aiActions.length === 0 ? (
                <div className="py-20 text-center space-y-4">
                  <div className="w-16 h-16 bg-white/5 rounded-full flex items-center justify-center mx-auto">
                    <Activity className="w-8 h-8 text-white/20" />
                  </div>
                  <p className="text-sm font-medium text-white/40 italic">No specific actions or reminders identified yet. Click "Refresh Analysis" to start.</p>
                </div>
              ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                  {aiActions
                    .filter(item => insightOwnerFilter === "All Owners" || item.leadOwner === insightOwnerFilter)
                    .map((item, idx) => (
                      <motion.div 
                        key={idx}
                        initial={{ opacity: 0, y: 20 }}
                        animate={{ opacity: 1, y: 0 }}
                        transition={{ delay: idx * 0.05 }}
                        className="p-6 bg-white/5 border border-white/5 rounded-2xl hover:bg-white/[0.08] transition-all group flex flex-col"
                      >
                        <div className="flex justify-between items-start mb-4">
                          <div className="flex items-center gap-3">
                            <div className={cn(
                              "w-2 h-2 rounded-full",
                              item.priority === 'high' ? "bg-rose-500 shadow-[0_0_10px_rgba(244,63,94,0.5)]" :
                              item.priority === 'medium' ? "bg-amber-500" : "bg-emerald-500"
                            )} />
                            <h4 className="text-xs font-black text-brand-accent uppercase tracking-widest leading-none">{item.organisation}</h4>
                          </div>
                          <span className={cn(
                            "text-[8px] font-black uppercase tracking-[0.2em] px-2 py-1 rounded border",
                            item.priority === 'high' ? "bg-rose-500/20 text-rose-400 border-rose-500/30" :
                            item.priority === 'medium' ? "bg-amber-500/20 text-amber-400 border-amber-500/30" :
                            "bg-emerald-500/20 text-emerald-400 border-emerald-500/30"
                          )}>
                            {item.priority}
                          </span>
                        </div>
                        
                        <p className="text-sm font-bold text-white mb-2 leading-snug">{item.action}</p>
                        <p className="text-xs font-medium text-white/40 leading-relaxed italic mb-6">{item.context}</p>
                        
                        <div className="mt-auto pt-4 border-t border-white/5 flex justify-between items-center">
                          <div className="flex items-center gap-2">
                            <UserCircle className="w-3 h-3 text-white/20" />
                            <span className="text-[9px] font-bold text-white/30 uppercase tracking-widest truncate max-w-[120px]">{item.leadOwner}</span>
                          </div>
                          <button 
                            onClick={() => {
                              const lead = leads.find(l => l.id === item.leadId);
                              if (lead) setViewingLead(lead);
                            }}
                            className="flex items-center gap-2 text-[10px] font-black uppercase tracking-widest text-white/30 group-hover:text-white transition-colors"
                          >
                            View Lead <ArrowUpRight className="w-3 h-3" />
                          </button>
                        </div>
                      </motion.div>
                    ))}
                </div>
              )}
            </div>
            
            <div className="mt-auto p-8 border-t border-white/5 bg-white/5 flex justify-center">
              <p className="text-[9px] font-bold text-white/20 uppercase tracking-[0.2em]">Insights are saved to the database for all team members after each analysis</p>
            </div>
          </motion.div>
        ) : viewMode === "board" ? (
          <DndContext 
            sensors={sensors}
            collisionDetection={closestCorners}
            onDragStart={handleDragStart}
            onDragEnd={handleDragEnd}
          >
            <motion.div 
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="flex gap-4 pb-12 min-w-max"
            >
              {boardStatuses.filter(s => selectedStatuses.length === 0 || selectedStatuses.includes(s)).map(status => (
                <div 
                  key={status} 
                  className="flex-shrink-0 flex flex-col gap-4"
                  style={{ width: `${columnWidths[status] || 320}px` }}
                >
                  <div className="flex items-center justify-between px-2">
                    <h4 className={cn(
                      "text-[10px] font-bold uppercase tracking-[0.2em] px-3 py-1.5 rounded-full relative group/status",
                      status.toLowerCase().includes("approved for future project") ? "bg-green-100 text-green-700" :
                      status.toLowerCase().includes("strong potential") ? "bg-blue-100 text-blue-700" :
                      status.toLowerCase().includes("under assessment") ? "bg-amber-100 text-amber-700" :
                      status.toLowerCase() === "assessed" ? "bg-emerald-100 text-emerald-700 border border-emerald-200" :
                      status.toLowerCase().includes("not suitable") || status.toLowerCase().includes("not interested") ? "bg-red-100 text-red-700" :
                      "bg-slate-100 text-slate-700"
                    )}>
                      {status}
                    </h4>
                    <span className="text-[10px] font-black text-brand-muted bg-white border border-brand-border w-5 h-5 flex items-center justify-center rounded-full">
                      {sortedLeads.filter(l => (l.leadStatus || "").toLowerCase().includes((status || "").toLowerCase())).length}
                    </span>
                  </div>
                  <DroppableStatusColumn status={status}>
                    {sortedLeads.filter(l => (l.leadStatus || "").toLowerCase().includes((status || "").toLowerCase())).map(lead => (
                      <DraggableLeadCard 
                        key={lead.id} 
                        lead={lead}
                        isAdminOrBoard={isAdminOrBoard}
                        canModifyStatus={canModifyStatus}
                        canAddLead={canAddLead}
                        canDeleteLead={canDeleteLead}
                        setViewingLead={setViewingLead}
                        handleDuplicate={handleDuplicate}
                        handleDelete={handleDelete}
                        onSelectResearch={onSelectResearch}
                        boardStatuses={allowedStatusesForRole}
                        notify={notify}
                        db={db}
                        handleFirestoreError={handleFirestoreError}
                      />
                    ))}
                    <button 
                      onClick={() => {
                        if (isAssessorRole && !isAdminOrBoard && (assessorPerms.disallowedStatuses || []).includes(status)) {
                          notify(`Only DCM board can add leads directly to "${status}"`, "error");
                          return;
                        }
                        setEditingLead({ organisation: "", leadStatus: status });
                        setIsEditing(true);
                      }}
                      className={cn(
                        "w-full flex items-center justify-center gap-2 py-3 border-2 border-dashed border-brand-border rounded-xl transition-all font-bold text-[10px] uppercase tracking-widest mt-2",
                        (isAssessorRole && !isAdminOrBoard && (assessorPerms.disallowedStatuses || []).includes(status))
                          ? "opacity-50 cursor-not-allowed text-brand-muted/40"
                          : "text-brand-muted hover:border-brand-accent/30 hover:text-brand-accent"
                      )}
                    >
                      <Plus className="w-3.5 h-3.5" /> Add New
                    </button>
                  </DroppableStatusColumn>
                </div>
              ))}
            </motion.div>
            <DragOverlay>
              {activeLead ? (
                <div className="rotate-3 shadow-2xl opacity-80 cursor-grabbing w-[300px]">
                  <LeadCardContent 
                    lead={activeLead}
                    isAdminOrBoard={isAdminOrBoard}
                    canModifyStatus={canModifyStatus}
                    canAddLead={canAddLead}
                    canDeleteLead={canDeleteLead}
                    setViewingLead={setViewingLead}
                    handleDuplicate={handleDuplicate}
                    handleDelete={handleDelete}
                    onSelectResearch={onSelectResearch}
                    boardStatuses={allowedStatusesForRole}
                    notify={notify}
                    db={db}
                    handleFirestoreError={handleFirestoreError}
                  />
                </div>
              ) : null}
            </DragOverlay>
          </DndContext>
        ) : (
          <motion.div 
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            className="bg-white border border-brand-border rounded-2xl overflow-hidden card-shadow"
          >
            <div className="bg-brand-bg/30 border-b border-brand-border px-6 py-3 flex gap-4 overflow-x-auto no-scrollbar">
              <button 
                onClick={() => setSelectedStatuses([])}
                className={cn(
                  "flex-none px-4 py-1.5 rounded-full text-[9px] font-black uppercase tracking-widest transition-all border",
                  selectedStatuses.length === 0 ? "bg-brand-primary text-white border-brand-primary" : "bg-white text-brand-muted border-brand-border hover:border-brand-accent"
                )}
              >
                All ({leads.length})
              </button>
              {boardStatuses.map(s => (
                <button 
                  key={s}
                  onClick={() => {
                    if (selectedStatuses.includes(s)) {
                      setSelectedStatuses(selectedStatuses.filter(item => item !== s));
                    } else {
                      setSelectedStatuses([...selectedStatuses, s]);
                    }
                  }}
                  className={cn(
                    "flex-none flex items-center gap-2 px-4 py-1.5 rounded-full text-[9px] font-black uppercase tracking-widest transition-all border",
                    selectedStatuses.includes(s) ? "bg-brand-primary text-white border-brand-primary" : "bg-white text-brand-muted border-brand-border hover:border-brand-accent"
                  )}
                >
                  {s} 
                  <span className={cn(
                    "px-1.5 rounded-full text-[8px]",
                    selectedStatuses.includes(s) ? "bg-white/20 text-white" : "bg-brand-bg text-brand-muted"
                  )}>
                    {leads.filter(l => (l.leadStatus || "").toLowerCase() === (s || "").toLowerCase()).length}
                  </span>
                </button>
              ))}
            </div>
            <div className="overflow-x-visible">
              <table className="min-w-full text-left border-collapse">
                <thead>
                  <tr className="bg-brand-bg/50 border-b border-brand-border">
                    {columnDefs.map(col => (
                      <th 
                        key={col.id} 
                        style={{ width: `${col.width}px` }}
                        className="px-6 py-4 text-[10px] font-black text-brand-primary uppercase tracking-[0.2em]"
                      >
                        {col.label}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-brand-bg">
                  {sortedLeads.map(lead => (
                    <tr key={lead.id} className="hover:bg-brand-bg/30 transition-colors group">
                      {columnDefs.map(col => (
                        <td key={col.id} className="px-6 py-4 overflow-hidden">
                          {col.id === 'organisation' && (
                            <div className="flex items-center gap-3">
                              <div className="overflow-hidden flex-1">
                                <div className="flex items-center gap-2 group/org min-w-0">
                                  <EditableTableCell 
                                    value={lead.organisation} 
                                    onSave={(val) => handleTableUpdate(lead.id, 'organisation', val)}
                                    className="text-sm font-black text-brand-primary"
                                    disabled={!canEditLead}
                                  />
                                  {lead.leadStatus === "Assessed" && (
                                    lead.approved ? (
                                      <div className="p-0.5 bg-emerald-500 rounded flex-shrink-0" title="Approved for Future Project">
                                        <CheckCircle2 className="w-2.5 h-2.5 text-white" />
                                      </div>
                                    ) : (
                                      <div className="p-0.5 bg-rose-500 rounded flex-shrink-0" title="Needs Board Approval">
                                        <ShieldAlert className="w-2.5 h-2.5 text-white" />
                                      </div>
                                    )
                                  )}
                                  <button onClick={() => setViewingLead(lead)} className="p-1 opacity-0 group-hover/org:opacity-100 hover:text-brand-accent focus:opacity-100 outline-none">
                                    <ExternalLink className="w-3 h-3" />
                                  </button>
                                  {lead.isVerified && <ShieldCheck className="w-3.5 h-3.5 text-emerald-500 shrink-0" />}
                                  {lead.approved && <CheckCircle2 className="w-3.5 h-3.5 text-emerald-500 shrink-0" />}
                                  {lead.isPaidEngagement && <DollarSign className="w-3.5 h-3.5 text-brand-accent shrink-0" />}
                                  {lead.needsFurtherAssessment && <AlertCircle className="w-3.5 h-3.5 text-amber-500 shrink-0" />}
                                </div>
                                <EditableTableCell 
                                  value={lead.website || ""} 
                                  onSave={(val) => handleTableUpdate(lead.id, 'website', val)}
                                  className="text-[10px] font-medium text-brand-muted"
                                  disabled={!canEditLead}
                                />
                              </div>
                            </div>
                          )}
                          {col.id === 'status' && (
                            <div className="relative group/status flex items-center">
                              {canModifyStatus ? (
                                <>
                                  <select 
                                    className="absolute inset-0 opacity-0 cursor-pointer z-10"
                                    value={lead.leadStatus}
                                    onChange={(e) => {
                                      const newStatus = e.target.value;
                                      setDoc(doc(db, "leads", lead.id), { leadStatus: newStatus, updatedAt: new Date().toISOString() }, { merge: true })
                                        .then(() => notify("Status updated", "success"))
                                        .catch((err) => {
                                          notify("Update failed", "error");
                                          handleFirestoreError(err, OperationType.UPDATE, "leads");
                                        });
                                    }}
                                  >
                                    {boardStatuses.map(s => <option key={s} value={s}>{s}</option>)}
                                  </select>
                                  <span className={cn(
                                    "text-[8px] font-black uppercase tracking-widest px-2.5 py-1 rounded-full whitespace-nowrap transition-all border border-transparent group-hover/status:border-brand-accent group-hover/status:text-brand-accent",
                                    (lead.leadStatus || "").toLowerCase().includes("approved for future project") ? "bg-green-100 text-green-700" :
                                    (lead.leadStatus || "").toLowerCase().includes("strong potential") ? "bg-blue-100 text-blue-700" :
                                    (lead.leadStatus || "").toLowerCase().includes("under assessment") ? "bg-amber-100 text-amber-700" :
                                    (lead.leadStatus || "").toLowerCase() === "assessed" ? "bg-emerald-100 text-emerald-700 border border-emerald-200" :
                                    (lead.leadStatus || "").toLowerCase().includes("not suitable") || (lead.leadStatus || "").toLowerCase().includes("not interested") ? "bg-red-100 text-red-700" :
                                    "bg-slate-100 text-slate-700"
                                  )}>
                                    {lead.leadStatus}
                                  </span>
                                </>
                              ) : (
                                <span className={cn(
                                  "text-[8px] font-black uppercase tracking-widest px-2.5 py-1 rounded-full whitespace-nowrap opacity-60",
                                  (lead.leadStatus || "").toLowerCase().includes("approved for future project") ? "bg-green-50 text-green-600" :
                                  (lead.leadStatus || "").toLowerCase().includes("strong potential") ? "bg-blue-50 text-blue-600" :
                                  (lead.leadStatus || "").toLowerCase().includes("under assessment") ? "bg-amber-50 text-amber-600" :
                                  (lead.leadStatus || "").toLowerCase() === "assessed" ? "bg-emerald-50 text-emerald-600" :
                                  "bg-slate-50 text-slate-600"
                                )}>
                                  {lead.leadStatus}
                                </span>
                              )}
                            </div>
                          )}
                          {col.id === 'owner' && <EditableTableCell value={lead.leadOwner || '-'} options={assignableOwnerOptions} type="select" onSave={(val) => handleTableUpdate(lead.id, 'leadOwner', val)} disabled={!canEditLead} />}
                          {col.id === 'type' && <EditableTableCell value={lead.leadType || '-'} options={leadTypes} onSave={(val) => handleTableUpdate(lead.id, 'leadType', val)} disabled={!canEditLead} />}
                          {col.id === 'lastContact' && <EditableTableCell value={lead.lastContact || '-'} type="date" format={settings?.displayConfig?.dateFormat} onSave={(val) => handleTableUpdate(lead.id, 'lastContact', val)} disabled={!canEditLead} />}
                          {col.id === 'potentialDate' && <EditableTableCell value={lead.projectPotentialDate || '-'} type="date" format={settings?.displayConfig?.dateFormat} onSave={(val) => handleTableUpdate(lead.id, 'projectPotentialDate', val)} disabled={!canEditLead} />}
                          {col.id === 'actions' && (
                            <div className="flex justify-end gap-2 pr-2 opacity-0 group-hover:opacity-100 transition-opacity">
                              <button 
                                onClick={() => handleDuplicate(lead)}
                                className="p-2 text-brand-muted hover:text-brand-accent hover:bg-brand-bg rounded-lg transition-all"
                                title="Duplicate Lead"
                              >
                                <Copy className="w-4 h-4" />
                              </button>
                              <button 
                                onClick={() => {
                                  setViewingLead(lead);
                                }}
                                className="p-2 text-brand-muted hover:text-brand-accent hover:bg-brand-bg rounded-lg transition-all"
                                title="Edit Lead"
                              >
                                <Edit3 className="w-4 h-4" />
                              </button>
                              {isAdminOrBoard && (
                                <button 
                                  onClick={() => handleDelete(lead.id)}
                                  className="p-2 text-brand-muted hover:text-red-500 hover:bg-red-50 rounded-lg transition-all"
                                  title="Delete Lead"
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              )}
                            </div>
                          )}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
              {filteredLeads.length === 0 && (
                <div className="p-12 text-center">
                  <p className="text-sm text-brand-muted font-medium italic">No leads found matching your criteria.</p>
                </div>
              )}
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {viewingLead && (
          <div className="fixed inset-0 bg-brand-primary/60 backdrop-blur-sm z-50 flex justify-end">
            <motion.div 
              initial={{ x: "100%" }}
              animate={{ x: 0 }}
              exit={{ x: "100%" }}
              transition={{ type: "spring", damping: 25, stiffness: 200 }}
              className="w-full max-w-xl bg-[#1A1A1A] text-white h-screen shadow-2xl overflow-y-auto flex flex-col"
            >
              <div className="p-8 border-b border-white/5 flex justify-between items-start sticky top-0 bg-[#1A1A1A] z-10">
                <div className="flex-1 min-w-0 pr-4">
                  {isEditingName ? (
                    <input 
                      autoFocus
                      className="text-3xl font-black tracking-tight bg-white/5 border border-white/10 rounded-lg px-2 py-1 outline-none focus:ring-1 focus:ring-brand-accent w-full text-white"
                      value={tempName}
                      onChange={(e) => setTempName(e.target.value)}
                      onBlur={() => {
                        handleUpdateField('organisation', tempName);
                        setIsEditingName(false);
                      }}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          handleUpdateField('organisation', tempName);
                          setIsEditingName(false);
                        }
                        if (e.key === 'Escape') {
                          setTempName(viewingLead.organisation || "");
                          setIsEditingName(false);
                        }
                      }}
                    />
                  ) : (
                    <div className={cn("flex items-center gap-3 group/name", canEditLead ? "cursor-pointer" : "cursor-default")} onClick={() => canEditLead && setIsEditingName(true)}>
                      <h2 className="text-3xl font-black tracking-tight truncate">{viewingLead.organisation}</h2>
                      {viewingLead.isVerified && <ShieldCheck className="w-6 h-6 text-emerald-500 shrink-0" />}
                    </div>
                  )}
                  <div className="flex items-center gap-3 mt-2">
                    <p 
                      className="text-[10px] font-black uppercase tracking-[0.2em] text-brand-accent cursor-pointer hover:underline"
                      onClick={() => handleUpdateField('leadStatus', viewingLead.leadStatus)}
                    >
                      {viewingLead.leadStatus}
                    </p>
                    {viewingLead.assessedBy && !viewingLead.needsFurtherAssessment && (
                      <>
                        <span className="w-1 h-1 rounded-full bg-white/20" />
                        <div className="flex items-center gap-1.5 opacity-80">
                          <CheckCircle2 className="w-2.5 h-2.5 text-emerald-500" />
                          <div className="text-[10px] font-bold text-white uppercase tracking-widest">
                            {viewingLead.approved ? (
                              <div className="flex flex-col">
                                <span className="block">Approved (Board)</span>
                                <span className="block text-[9px] mt-0.5 font-medium opacity-80 leading-none">
                                  {viewingLead.approvedBy?.split('@')[0] || 'Admin'} • {viewingLead.approvedAt ? new Date(viewingLead.approvedAt).toLocaleDateString() : '?'}
                                </span>
                              </div>
                            ) : (
                              <div className="flex flex-col">
                                <span className={cn(
                                  "block",
                                  (viewingLead.leadStatus || "").toLowerCase().includes("under assessment") ? "text-amber-400" : 
                                  (viewingLead.leadStatus || "").toLowerCase() === "assessed" ? "text-blue-400" : ""
                                )}>
                                  {(viewingLead.leadStatus || "").toLowerCase().includes("under assessment") ? "Further Assessment" : "Awaiting Approval"}
                                </span>
                                <span className="block text-[9px] mt-0.5 font-medium opacity-80 leading-none">
                                  {(viewingLead.assessedBy || "").split('@')[0]} ({viewingLead.assessedAt ? new Date(viewingLead.assessedAt).toLocaleDateString() : "Pending"})
                                </span>
                              </div>
                            )}
                          </div>
                        </div>
                      </>
                    )}
                    <span className="w-1 h-1 rounded-full bg-white/20" />
                    <button 
                      onClick={handleSuggestFollowUp}
                      disabled={isSuggestingEmail}
                      className="text-[10px] font-black uppercase tracking-[0.2em] text-white/40 hover:text-brand-accent flex items-center gap-1.5 transition-colors disabled:opacity-50"
                    >
                      {isSuggestingEmail ? <Loader2 className="w-2.5 h-2.5 animate-spin" /> : <Sparkles className="w-2.5 h-2.5 text-brand-accent" />}
                      Suggest Follow-up
                    </button>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button onClick={() => { setViewingLead(null); setSuggestedEmail(null); }} className="p-2 hover:bg-white/5 rounded-full text-white/40 hover:text-white transition-all">
                    <X className="w-6 h-6" />
                  </button>
                </div>
              </div>

              <AnimatePresence>
                {suggestedEmail && (
                  <motion.div 
                    initial={{ opacity: 0, height: 0 }}
                    animate={{ opacity: 1, height: 'auto' }}
                    exit={{ opacity: 0, height: 0 }}
                    className="bg-brand-accent/10 border-b border-brand-accent/20 overflow-hidden"
                  >
                    <div className="p-8 space-y-4">
                      <div className="flex justify-between items-center">
                        <div className="flex items-center gap-2">
                          <Sparkles className="w-4 h-4 text-brand-accent" />
                          <h4 className="text-xs font-black uppercase tracking-widest text-brand-accent">AI Follow-up Draft</h4>
                        </div>
                        <div className="flex gap-2">
                          <button 
                            onClick={() => {
                              const subjectMatch = suggestedEmail.match(/Subject: (.*)/);
                              const subject = subjectMatch ? encodeURIComponent(subjectMatch[1]) : encodeURIComponent(`Follow up: ${viewingLead.organisation}`);
                              const body = encodeURIComponent(suggestedEmail.replace(/Subject: .*\n/, "").trim());
                              window.location.href = `mailto:${viewingLead.email || ""}?subject=${subject}&body=${body}`;
                            }}
                            className="bg-brand-accent text-white px-3 py-1 rounded text-[9px] font-black uppercase tracking-widest hover:scale-105 transition-all"
                          >
                            Open in Mail
                          </button>
                          <button 
                            onClick={() => {
                              navigator.clipboard.writeText(suggestedEmail);
                              alert("Draft copied!");
                            }}
                            className="bg-white/10 text-white px-3 py-1 rounded text-[9px] font-black uppercase tracking-widest hover:bg-white/20 transition-all border border-white/10"
                          >
                            Copy Draft
                          </button>
                          <button onClick={() => setSuggestedEmail(null)} className="text-white/40 hover:text-white text-[9px] font-black uppercase tracking-widest">
                            Close
                          </button>
                        </div>
                      </div>
                      <div className="bg-black/20 rounded-xl p-4 border border-white/5 max-h-60 overflow-y-auto">
                        <div className="prose prose-sm prose-invert max-w-none text-white/80 text-[10px] leading-relaxed whitespace-pre-wrap font-sans">
                          {suggestedEmail}
                        </div>
                      </div>
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>

              <div className="p-8 space-y-8 flex-1">
                {/* Financials & Engagement Section */}
                <div className="pt-8 border-t border-white/10 space-y-6">
                  <div className="flex items-center justify-between">
                    <h4 className="text-[10px] font-black uppercase tracking-[0.2em] text-amber-500">Financials & Engagement</h4>
                    <button 
                      onClick={() => handleUpdateField('isPaidEngagement', !viewingLead.isPaidEngagement)}
                      className={cn(
                        "px-3 py-1.5 rounded-lg text-[9px] font-black uppercase tracking-widest transition-all border",
                        viewingLead.isPaidEngagement 
                          ? "bg-amber-500 text-white border-amber-500" 
                          : "bg-white/5 text-white/40 border-white/10 hover:bg-white/10"
                      )}
                    >
                      {viewingLead.isPaidEngagement ? "Paid Engagement" : "Free Engagement"}
                    </button>
                  </div>

                  <div className={cn(
                    "bg-white/5 border rounded-2xl p-6 space-y-6 transition-all",
                    viewingLead.isPaidEngagement ? "border-amber-500/40 bg-amber-500/5 shadow-[0_0_20px_rgba(245,158,11,0.05)]" : "border-white/10"
                  )}>
                    <div className="grid grid-cols-2 gap-8">
                      <LeadDetailItem 
                        icon={DollarSign} 
                        label="Fees Paid ($CAD)" 
                        value={viewingLead.feesPaid} 
                        onSave={(val) => handleUpdateField('feesPaid', val)} 
                      />
                      <div className="space-y-4">
                        <div className="flex justify-between items-center">
                          <div className="flex items-center gap-2">
                            <Hash className="w-3.5 h-3.5 text-white/20" />
                            <span className="text-[10px] font-black uppercase tracking-[0.2em] text-white/20">Invoice Management</span>
                          </div>
                          {viewingLead.invoiceName && (
                            <span className="text-[8px] font-bold text-white/30 uppercase tracking-widest">
                              {new Date(viewingLead.invoiceUploadedAt).toLocaleDateString()}
                            </span>
                          )}
                        </div>
                        
                        {viewingLead.invoiceName ? (
                          <div className="flex items-center justify-between bg-black/20 rounded-xl p-3 border border-white/5">
                            <div className="flex items-center gap-3 min-w-0">
                              <Paperclip className="w-4 h-4 text-amber-500" />
                              <p className="text-[11px] font-bold text-white truncate">{viewingLead.invoiceName}</p>
                            </div>
                            <div className="flex items-center gap-2">
                              <a 
                                href={viewingLead.invoiceData} 
                                download={viewingLead.invoiceName || 'invoice.pdf'}
                                className="p-2 hover:bg-white/10 rounded-lg text-white/40 hover:text-white transition-all"
                                title="Download Invoice"
                              >
                                <Download className="w-3.5 h-3.5" />
                              </a>
                            </div>
                          </div>
                        ) : (
                          <div className="py-4 text-center border-2 border-dashed border-white/5 rounded-xl">
                            <p className="text-[10px] italic text-white/20">No invoice on file</p>
                          </div>
                        )}
                        
                        <label className={cn(
                          "w-full flex items-center justify-center gap-2 py-2.5 bg-white/5 border border-white/10 rounded-xl text-[9px] font-black uppercase tracking-widest transition-all text-white/60",
                          canEditLead ? "cursor-pointer hover:bg-white/10" : "cursor-not-allowed opacity-50"
                        )}>
                          <UploadCloud className="w-3.5 h-3.5 text-amber-500" />
                          {viewingLead.invoiceName ? "Update Invoice" : "Upload Invoice"}
                          <input 
                            type="file" 
                            className="hidden" 
                            disabled={!canEditLead}
                            onChange={async (e) => {
                              const file = e.target.files?.[0];
                              if (!file) return;
                              if (file.size > 800000) {
                                notify("File too large. Max 800KB.", "error");
                                return;
                              }
                              const reader = new FileReader();
                              reader.onload = async (event) => {
                                const base64Data = event.target?.result as string;
                                const updates = {
                                  invoiceData: base64Data,
                                  invoiceName: file.name,
                                  invoiceUploadedAt: new Date().toISOString()
                                };
                                try {
                                  await setDoc(doc(db, "leads", viewingLead.id), { ...updates, updatedAt: new Date().toISOString() }, { merge: true });
                                  setViewingLead({ ...viewingLead, ...updates });
                                  notify("Invoice uploaded", "success");
                                } catch (err) {
                                  notify("Upload failed", "error");
                                }
                              };
                              reader.readAsDataURL(file);
                            }} 
                          />
                        </label>
                      </div>
                    </div>
                  </div>
                </div>
                  {viewingLead.leadStatus === "Assessed" && !viewingLead.approved && !viewingLead.needsFurtherAssessment && (
                    <div className="bg-blue-500/10 border border-blue-500/20 rounded-2xl p-6 mb-6 shadow-[0_0_20px_rgba(59,130,246,0.05)]">
                      <div className="flex items-center gap-3 text-blue-500">
                        <div className="w-10 h-10 rounded-xl bg-blue-500 flex items-center justify-center text-white shadow-lg shadow-blue-500/20">
                          <Clock className="w-5 h-5" />
                        </div>
                        <div>
                          <p className="text-xs font-black uppercase tracking-widest leading-none">Awaiting Final Approval</p>
                          <p className="text-[9px] font-bold text-white/40 uppercase tracking-widest mt-1">This lead is assessed and waiting for DCM Board approval.</p>
                        </div>
                      </div>
                    </div>
                  )}
                  {viewingLead.needsFurtherAssessment && (
                  <div className="bg-amber-500/10 border border-amber-500/20 rounded-2xl p-6 mb-6 shadow-[0_0_20px_rgba(245,158,11,0.05)]">
                    <div className="flex items-center gap-3 text-amber-500">
                      <div className="w-10 h-10 rounded-xl bg-amber-500 flex items-center justify-center text-white shadow-lg shadow-amber-500/20">
                        <AlertCircle className="w-5 h-5" />
                      </div>
                      <div>
                        <p className="text-xs font-black uppercase tracking-widest leading-none">Intelligence Required</p>
                        <p className="text-[9px] font-bold text-white/40 uppercase tracking-widest mt-1">The board has requested further assessment before approval.</p>
                      </div>
                    </div>
                  </div>
                )}
                <LeadDetailItem icon={Activity} label="Activity" value={viewingLead.activity} onSave={(val) => handleUpdateField('activity', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={MapPin} label="City" value={viewingLead.city} onSave={(val) => handleUpdateField('city', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={MessageSquare} label="Comments" value={viewingLead.comments} onSave={(val) => handleUpdateField('comments', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={UserIcon} label="Contact Name" value={viewingLead.contactName} onSave={(val) => handleUpdateField('contactName', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={Globe} label="Country" value={viewingLead.country} onSave={(val) => handleUpdateField('country', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={AtSign} label="Email" value={viewingLead.email} onSave={(val) => handleUpdateField('email', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={Calendar} label="Last Contact" value={viewingLead.lastContact} type="date" onSave={(val) => handleUpdateField('lastContact', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={Target} label="Lead status" value={viewingLead.leadStatus} isBadge options={allowedStatusesForRole} onSave={(val) => handleUpdateField('leadStatus', val)} disabled={!canModifyStatus} />
                <LeadDetailItem icon={Target} label="Lead Source" value={viewingLead.leadSource} isBadge options={leadSources} onSave={(val) => handleUpdateField('leadSource', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={Languages} label="Project Languages" value={viewingLead.projectLanguages} onSave={(val) => handleUpdateField('projectLanguages', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={CalendarDays} label="Project Potential Date" value={viewingLead.projectPotentialDate} type="date" onSave={(val) => handleUpdateField('projectPotentialDate', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={Users} label="Referred by" value={viewingLead.referredBy} onSave={(val) => handleUpdateField('referredBy', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={UserCircle} label="Lead owner" value={viewingLead.leadOwner} options={assignableOwnerOptions} onSave={(val) => handleUpdateField('leadOwner', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={Link} label="Website" value={viewingLead.website} isLink onSave={(val) => handleUpdateField('website', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={ExternalLink} label="Direct link" value={viewingLead.directLink} isLink onSave={(val) => handleUpdateField('directLink', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={FastForward} label="What's next?" value={viewingLead.whatsNext} onSave={(val) => handleUpdateField('whatsNext', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={Heart} label="Donation" value={viewingLead.donation} onSave={(val) => handleUpdateField('donation', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={CheckCircle2} label="Assessed By" value={viewingLead.assessedBy} onSave={(val) => handleUpdateField('assessedBy', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={CalendarDays} label="Assessed Date" value={viewingLead.assessedAt} type="date" onSave={(val) => handleUpdateField('assessedAt', val)} disabled={!canEditLead} />
                {viewingLead.approved && (
                  <>
                    <LeadDetailItem icon={ShieldCheck} label="Approved (Board)" value={viewingLead.approvedBy} onSave={(val) => handleUpdateField('approvedBy', val)} disabled={!isAdminOrBoard} />
                    <LeadDetailItem icon={CalendarDays} label="Approval Date" value={viewingLead.approvedAt} type="date" onSave={(val) => handleUpdateField('approvedAt', val)} disabled={!isAdminOrBoard} />
                  </>
                )}
                <LeadDetailItem icon={Tag} label="Lead Type" value={viewingLead.leadType} isBadge options={leadTypes} onSave={(val) => handleUpdateField('leadType', val)} disabled={!canEditLead} />
                <LeadDetailItem icon={Paperclip} label="Attachments" value={viewingLead.attachments} onSave={(val) => handleUpdateField('attachments', val)} disabled={!canEditLead} />

                <div className="pt-8 border-t border-white/10 space-y-6">
                  <h4 className="text-[10px] font-black uppercase tracking-[0.2em] text-blue-500">Research & Intelligence</h4>
                  <div className="space-y-6">
                    <LeadDetailItem icon={DollarSign} label="Financial Tier / Revenue" value={viewingLead.revenue} onSave={(val) => handleUpdateField('revenue', val)} disabled={!canEditLead} />
                    <LeadDetailItem icon={Fingerprint} label="EIN Number" value={viewingLead.ein} onSave={(val) => handleUpdateField('ein', val)} disabled={!canEditLead} />
                    <LeadDetailItem icon={FileText} label="DCM Brief Summary" value={viewingLead.briefSummary} onSave={(val) => handleUpdateField('briefSummary', val)} disabled={!canEditLead} />
                    <LeadDetailItem icon={Star} label="Charity Navigator Rating" value={viewingLead.charity_navigator_rating} onSave={(val) => handleUpdateField('charity_navigator_rating', val)} disabled={!canEditLead} />
                    <LeadDetailItem icon={Linkedin} label="LinkedIn Overview" value={viewingLead.linkedin_overview} onSave={(val) => handleUpdateField('linkedin_overview', val)} disabled={!canEditLead} />
                  </div>
                </div>

                {/* Board Approval Section */}
                {isAdminOrBoard && (
                  <div className="pt-8 border-t border-white/10 space-y-6">
                    <h4 className="text-[10px] font-black uppercase tracking-[0.2em] text-emerald-500 text-center">Board Approval Workflow</h4>
                    <div className={cn(
                      "bg-white/5 border rounded-2xl p-6 space-y-6 transition-all",
                      viewingLead.approved ? "border-emerald-500/40 bg-emerald-500/5 shadow-[0_0_20px_rgba(16,185,129,0.05)]" : 
                      viewingLead.needsFurtherAssessment ? "border-amber-500/40 bg-amber-500/5 shadow-[0_0_20px_rgba(245,158,11,0.05)]" : "border-white/10"
                    )}>
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-3">
                          <div className={cn(
                            "w-10 h-10 rounded-xl flex items-center justify-center transition-all",
                            viewingLead.approved ? "bg-emerald-500 text-white shadow-lg shadow-emerald-500/20" : 
                            viewingLead.needsFurtherAssessment ? "bg-amber-500 text-white shadow-lg shadow-amber-500/20" : "bg-white/10 text-white/40"
                          )}>
                            {viewingLead.needsFurtherAssessment ? <AlertCircle className="w-5 h-5" /> : <CheckCircle2 className={cn("w-5 h-5", viewingLead.approved ? "animate-pulse" : "")} />}
                          </div>
                          <div>
                            <p className="text-xs font-black uppercase tracking-widest text-white">
                              {viewingLead.needsFurtherAssessment ? "Assessment Flag Active" : "Board Approved"}
                            </p>
                            <p className="text-[9px] font-bold text-white/40 uppercase tracking-widest mt-1">
                              {viewingLead.approved 
                                ? `By ${viewingLead.approvedBy} on ${new Date(viewingLead.approvedAt).toLocaleDateString()}` 
                                : viewingLead.needsFurtherAssessment
                                  ? "Lead requires further deliberation"
                                  : "Mark this lead as approved by the board"}
                            </p>
                          </div>
                        </div>
                        <div className="flex gap-2">
                          {!viewingLead.approved && (
                            <>
                              {viewingLead.needsFurtherAssessment && (
                                <button 
                                  onClick={async () => {
                                    const updates: any = {
                                      needsFurtherAssessment: false,
                                      updatedAt: new Date().toISOString()
                                    };
                                    
                                    // Add a system note
                                    const newNote = {
                                      id: Math.random().toString(36).substr(2, 9),
                                      author: userProfile?.name || userProfile?.email?.split('@')[0] || "Board Member",
                                      text: "CLEARED: Needs Further Assessment flag removed.",
                                      date: new Date().toLocaleString()
                                    };
                                    updates.notes = [newNote, ...(viewingLead.notes || [])];

                                    try {
                                      await setDoc(doc(db, "leads", viewingLead.id), updates, { merge: true });
                                      setViewingLead({ ...viewingLead, ...updates });
                                      await logLeadEvent(viewingLead.id, "BOARD_FLAG_CLEARED", "Further assessment flag cleared by board");
                                      notify("Further assessment flag cleared.", "success");
                                    } catch (err) {
                                      handleFirestoreError(err, OperationType.UPDATE, "leads");
                                    }
                                  }}
                                  className="px-5 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest bg-blue-500/10 text-blue-500 border border-blue-500/20 hover:bg-blue-500/20 transition-all shadow-sm"
                                >
                                  Clear Flag
                                </button>
                              )}
                              <button 
                                onClick={async () => {
                                  const comment = approvalCommentLocal || "";
                                  if (!comment.trim()) {
                                    notify("Please provide a comment explaining what needs further assessment.", "error");
                                    return;
                                  }

                                  const updates: any = {
                                    leadStatus: "Under assessment",
                                    needsFurtherAssessment: true,
                                    approvalComment: comment,
                                    approved: false,
                                    approvedAt: null,
                                    approvedBy: null,
                                    updatedAt: new Date().toISOString()
                                  };

                                  // Add a system note
                                  const newNote = {
                                    id: Math.random().toString(36).substr(2, 9),
                                    author: userProfile?.name || userProfile?.email?.split('@')[0] || "Board Member",
                                    text: `NEEDS FURTHER ASSESSMENT: ${comment}`,
                                    date: new Date().toLocaleString()
                                  };
                                  updates.notes = [newNote, ...(viewingLead.notes || [])];

                                  try {
                                    await setDoc(doc(db, "leads", viewingLead.id), updates, { merge: true });
                                    await logLeadEvent(viewingLead.id, "BOARD_REVOKED", "Sent back for further assessment", { comment });
                                    setViewingLead({ ...viewingLead, ...updates });
                                    notify("Sent back for further assessment.", "success");
                                  } catch (err) {
                                    handleFirestoreError(err, OperationType.UPDATE, "leads");
                                  }
                                }}
                                className={cn(
                                  "px-5 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all shadow-sm",
                                  viewingLead.needsFurtherAssessment 
                                    ? "bg-amber-500 text-white border-amber-600 shadow-lg shadow-amber-500/20" 
                                    : "bg-amber-500/10 text-amber-500 border border-amber-500/20 hover:bg-amber-500/20"
                                )}
                              >
                                {viewingLead.needsFurtherAssessment ? "Update Assessment" : "Needs Further Assessment"}
                              </button>
                            </>
                          )}
                          <button 
                            onClick={async () => {
                              const newApproved = !viewingLead.approved;
                              const approvedAt = new Date().toISOString();
                              const approvedBy = userProfile?.email || "Board Member";
                              
                              const updates: any = {
                                approved: newApproved,
                                approvedAt: newApproved ? approvedAt : null,
                                approvedBy: newApproved ? approvedBy : null,
                                needsFurtherAssessment: false,
                                updatedAt: new Date().toISOString()
                              };
                              
                              if (newApproved) {
                                const targetStatus = boardStatuses.find(s => s.toLowerCase().includes("approved for future project")) || "Approved";
                                updates.leadStatus = targetStatus;
                              } else {
                                updates.leadStatus = "Under assessment";
                              }
                              
                              try {
                                await setDoc(doc(db, "leads", viewingLead.id), updates, { merge: true });
                                setViewingLead({ ...viewingLead, ...updates });
                                await logLeadEvent(viewingLead.id, newApproved ? "BOARD_APPROVED" : "BOARD_REVOKED", newApproved ? "Lead Approved by Board" : "Board Approval Revoked - Moved to Under Assessment");
                                notify(newApproved ? "Lead Approved!" : "Board Approval Revoked. Status reverted to Under Assessment.", "success");
                              } catch (err) {
                                handleFirestoreError(err, OperationType.UPDATE, "leads");
                              }
                            }}
                            className={cn(
                              "px-5 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all shadow-sm",
                              viewingLead.approved 
                                ? "bg-emerald-500 text-white hover:bg-emerald-600" 
                                : "bg-white/10 text-white/60 hover:bg-white/20 border border-white/5"
                            )}
                          >
                            {viewingLead.approved ? "Revoke" : "Approve"}
                          </button>
                        </div>
                      </div>
                      
                      <div className="space-y-3">
                        <div className="flex justify-between items-center ml-1">
                          <label className="text-[10px] font-black uppercase tracking-widest text-white/30">Workflow Comment</label>
                          <button 
                            onClick={() => handleUpdateField('approvalComment', approvalCommentLocal)}
                            className="text-[9px] font-black text-brand-accent uppercase tracking-widest hover:underline"
                          >
                            Save Comment
                          </button>
                        </div>
                        <textarea 
                          className="w-full bg-black/20 border border-white/10 rounded-2xl px-4 py-4 text-xs text-white font-medium focus:ring-2 focus:ring-emerald-500/50 outline-none h-28 transition-all resize-none placeholder:text-white/10"
                          placeholder="Add a board comment regarding this approval or reason for further assessment..."
                          value={approvalCommentLocal}
                          onChange={(e) => setApprovalCommentLocal(e.target.value)}
                        />
                      </div>
                    </div>
                  </div>
                )}

                <div className="pt-8 border-t border-white/10 space-y-6">
                  <h4 className="text-[10px] font-black uppercase tracking-[0.2em] text-brand-accent">Assessment Pack (PDF)</h4>
                  <div className="bg-white/5 border border-white/10 rounded-2xl p-6 space-y-4">
                    {viewingLead.assessmentPdf ? (
                      <div className="flex items-center justify-between gap-4">
                        <div className="flex items-center gap-3 min-w-0">
                          <div className="p-2 bg-emerald-500/20 rounded-lg flex-shrink-0">
                            <FileText className="w-5 h-5 text-emerald-500" />
                          </div>
                          <div className="min-w-0">
                            <p className="text-xs font-bold text-white truncate">{viewingLead.assessmentPdfName || "assessment_pack.pdf"}</p>
                            <p className="text-[9px] font-medium text-white/40 uppercase tracking-widest mt-0.5 truncate">
                              Uploaded by {viewingLead.assessmentUploadedBy || "Unknown"} • {viewingLead.assessmentUploadedAt ? new Date(viewingLead.assessmentUploadedAt).toLocaleString() : "Recently"}
                            </p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 flex-shrink-0">
                          <a 
                            href={viewingLead.assessmentPdf?.startsWith('data:') ? viewingLead.assessmentPdf : '#'}
                            onClick={(e) => {
                              if (!viewingLead.assessmentPdf?.startsWith('data:')) {
                                e.preventDefault();
                                notify("Full PDF data not found. This was likely uploaded as a reference filename only.", "error");
                              }
                            }}
                            download={viewingLead.assessmentPdfName || "assessment_pack.pdf"}
                            className="px-3 py-2 bg-emerald-500/10 hover:bg-emerald-500/20 rounded-lg text-emerald-500 transition-all flex items-center gap-2 text-[9px] font-black uppercase tracking-widest border border-emerald-500/20"
                            title="Download PDF"
                          >
                            <Download className="w-3.5 h-3.5" />
                            <span className="hidden sm:inline">Download</span>
                          </a>
                          <button 
                            onClick={() => {
                              if (viewingLead.assessmentPdf?.startsWith('data:')) {
                                const win = window.open();
                                if (win) {
                                  win.document.write(`<iframe src="${viewingLead.assessmentPdf}" frameborder="0" style="border:0; top:0px; left:0px; bottom:0px; right:0px; width:100%; height:100%;" allowfullscreen></iframe>`);
                                }
                              } else {
                                notify("Full file data missing. Cannot view in browser.", "error");
                              }
                            }}
                            className="p-2 hover:bg-white/10 rounded-lg text-white/40 hover:text-white transition-all flex items-center gap-2 text-[9px] font-black uppercase tracking-widest"
                            title="View PDF"
                          >
                            <ExternalLink className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </div>
                    ) : (
                      <div className="text-center py-6 border-2 border-dashed border-white/10 rounded-xl">
                        <p className="text-xs font-medium text-white/40 italic mb-4">No assessment pack uploaded yet.</p>
                      </div>
                    )}
                    
                    <label className={cn(
                      "w-full flex items-center justify-center gap-2 py-3 bg-brand-primary border border-white/10 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all",
                      canEditLead ? "cursor-pointer hover:bg-white/5" : "cursor-not-allowed opacity-50"
                    )}>
                      <UploadCloud className="w-4 h-4 text-brand-accent" />
                      {viewingLead.assessmentPdf ? "Update Assessment PDF" : "Upload Assessment PDF"}
                      <input 
                        type="file" 
                        className="hidden" 
                        accept="application/pdf"
                        disabled={!canEditLead}
                        onChange={(e) => {
                          const file = e.target.files?.[0];
                          if (!file) return;
                          
                          // Increased limit for direct storage (Firestore Document limit is 1MB total)
                          // 950KB allows space for secondary fields
                          if (file.size > 950000) { 
                            notify("File too large. Max size is 950KB for direct storage. Try generating a more compressed version.", "error");
                            return;
                          }

                          const reader = new FileReader();
                          reader.onload = async (event) => {
                            const base64Data = event.target?.result as string;
                            const uploader = userProfile?.name || auth.currentUser?.email || "Assessor";
                            const timestamp = new Date().toISOString();
                            
                            try {
                              await setDoc(doc(db, "leads", viewingLead.id), {
                                assessmentPdf: base64Data, 
                                assessmentPdfName: file.name,
                                assessmentUploadedBy: uploader,
                                assessmentUploadedAt: timestamp,
                                leadStatus: "Assessed",
                                updatedAt: timestamp
                              }, { merge: true });
                              
                              await logLeadEvent(viewingLead.id, "ASSESSMENT_UPLOAD", `Assessment pack uploaded: ${file.name}`, { filename: file.name });
                              
                              setViewingLead({
                                ...viewingLead,
                                assessmentPdf: base64Data,
                                assessmentPdfName: file.name,
                                assessmentUploadedBy: uploader,
                                assessmentUploadedAt: timestamp,
                                leadStatus: "Assessed"
                              });
                              
                              notify("Assessment pack uploaded. Status updated to Assessed.", "success");
                            } catch (err) {
                              console.error("Upload error:", err);
                              notify("Failed to upload assessment.", "error");
                            }
                          };
                          reader.onerror = () => notify("Failed to read file", "error");
                          reader.readAsDataURL(file);
                        }}
                      />
                    </label>
                  </div>
                </div>

                <div className="pt-8 border-t border-white/10 space-y-6 pb-12">
                  <h4 className="text-[10px] font-black uppercase tracking-[0.2em] text-white/40">Timeline & Internal Notes</h4>
                  
                  <div className="space-y-6 relative before:absolute before:left-6 before:top-4 before:bottom-4 before:w-[2px] before:bg-white/5">
                    {(() => {
                      const mergedTimeline = [
                        ...(leadTimeline || []).map(e => ({ ...e, isEvent: true, sortDate: new Date(e.timestamp) })),
                        ...(viewingLead.notes || []).map(n => {
                           let d;
                           try { d = new Date(n.date); if(isNaN(d.getTime())) d = new Date(); } catch { d = new Date(); }
                           return { ...n, isNote: true, sortDate: d };
                        })
                      ].sort((a, b) => b.sortDate.getTime() - a.sortDate.getTime());

                      if (mergedTimeline.length === 0) {
                        return <p className="text-center py-12 text-xs italic text-white/20">No history recorded yet.</p>
                      }

                      return mergedTimeline.map((item: any, idx: number) => (
                        <div key={idx} className="relative pl-14 group">
                          <div className={cn(
                            "absolute left-4 top-1 w-4 h-4 rounded-full border-4 transition-all z-10",
                            item.isEvent 
                              ? (item.type?.includes("APPROVED") 
                                  ? "bg-emerald-500 border-emerald-500/20 shadow-[0_0_10px_rgba(16,185,129,0.3)]" 
                                  : item.type?.includes("REVOKED")
                                    ? "bg-amber-500 border-amber-500/20 shadow-[0_0_10px_rgba(245,158,11,0.3)]"
                                    : "bg-brand-accent border-brand-accent/20 shadow-[0_0_10px_rgba(255,54,142,0.3)]")
                              : "bg-blue-500 border-blue-500/20 shadow-[0_0_10px_rgba(59,130,246,0.3)]"
                          )} />
                          
                          <div className={cn(
                            "p-4 rounded-xl border space-y-2 transition-all hover:translate-x-1",
                            item.isEvent ? "bg-brand-bg/40 border-brand-accent/10" : "bg-white/5 border-white/5"
                          )}>
                            <div className="flex justify-between items-center text-[10px] font-black uppercase tracking-widest gap-4">
                              <span className={cn(
                                item.isEvent 
                                  ? (item.type?.includes("APPROVED") 
                                      ? "text-emerald-500" 
                                      : item.type?.includes("REVOKED")
                                        ? "text-amber-500"
                                        : "text-brand-accent")
                                  : "text-blue-400"
                              )}>
                                {item.isEvent ? item.type.replace(/_/g, ' ') : "INTERNAL NOTE"}
                              </span>
                              <span className="text-[9px] text-white/20 whitespace-nowrap">
                                {item.isEvent 
                                  ? new Date(item.timestamp).toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' }) 
                                  : item.date}
                              </span>
                            </div>
                            
                            <p className="text-xs font-bold text-white/80 leading-relaxed">
                              {item.isEvent ? item.description : item.text}
                            </p>
                            
                            <div className="flex items-center gap-2 pt-1 border-t border-white/5">
                              <div className="w-4 h-4 rounded-full bg-white/5 flex items-center justify-center">
                                <UserIcon className="w-2.5 h-2.5 text-white/40" />
                              </div>
                              <span className="text-[9px] font-bold text-white/30 truncate">
                                {item.author || item.authorEmail || "Unknown User"}
                              </span>
                            </div>
                          </div>
                        </div>
                      ));
                    })()}
                  </div>

                  <div className="space-y-3">
                    <textarea 
                      placeholder="Add an internal note or update..."
                      className="w-full bg-white/5 border border-white/10 rounded-xl p-4 text-sm font-medium outline-none focus:ring-1 focus:ring-brand-accent transition-all resize-none min-h-[100px] text-white"
                      value={newComment}
                      onChange={(e) => setNewComment(e.target.value)}
                    />
                    <button 
                      onClick={handleAddComment}
                      disabled={!newComment.trim()}
                      className="w-full py-3 bg-white/10 hover:bg-white/20 text-white rounded-xl font-bold text-xs uppercase tracking-widest transition-all disabled:opacity-50"
                    >
                      Add Comment
                    </button>
                  </div>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function HelpModal({ isOpen, onClose }: { isOpen: boolean, onClose: () => void }) {
  if (!isOpen) return null;

  const steps = [
    {
      title: "1. Start Assessment",
      description: "Begin by accessing leads with 'Under Assessment' status in DCM Leads or directly from your Workplace dashboard."
    },
    {
      title: "2. Deep Research",
      description: "Use the Deep Research Agent to uncover intelligence, LinkedIn presence, and financial data for the organization."
    },
    {
      title: "3. Scoring Workspace",
      description: "Use the Scoring tool to evaluate the lead. Review project briefs and data samples, and answer all validation questions."
    },
    {
      title: "4. Assessment Pack",
      description: "Generate the Assessment Pack using the gathered intelligence and scores. Refine the suggested recommendations for the board."
    },
    {
      title: "5. Final Submission",
      description: "Download the Assessment Pack PDF and upload it back to the Lead record. The status will update to 'Assessed' automatically."
    }
  ];

  return (
    <AnimatePresence>
      <motion.div 
        initial={{ opacity: 0 }} 
        animate={{ opacity: 1 }} 
        exit={{ opacity: 0 }}
        className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-brand-primary/80 backdrop-blur-sm"
        onClick={onClose}
      >
        <motion.div 
          initial={{ scale: 0.9, opacity: 0, y: 20 }}
          animate={{ scale: 1, opacity: 1, y: 0 }}
          exit={{ scale: 0.9, opacity: 0, y: 20 }}
          className="bg-white rounded-3xl w-full max-w-2xl overflow-hidden shadow-2xl"
          onClick={e => e.stopPropagation()}
        >
          <div className="p-8 border-b border-brand-border flex justify-between items-center bg-brand-bg">
            <div>
              <h3 className="text-xl font-black text-brand-primary uppercase tracking-tight">Assessment Process</h3>
              <p className="text-xs font-bold text-brand-muted uppercase tracking-widest mt-1">Guide for DCM Assessors</p>
            </div>
            <button onClick={onClose} className="p-2 hover:bg-brand-border rounded-full transition-all">
              <X className="w-5 h-5 text-brand-muted" />
            </button>
          </div>
          
          <div className="p-8 space-y-6 max-h-[70vh] overflow-y-auto">
            {steps.map((step, idx) => (
              <div key={idx} className="flex gap-6 items-start group">
                <div className="w-10 h-10 rounded-xl bg-brand-accent/10 border border-brand-accent/20 flex items-center justify-center flex-shrink-0 font-black text-brand-accent text-sm group-hover:scale-110 transition-transform">
                  {idx + 1}
                </div>
                <div className="pt-1">
                  <h4 className="text-sm font-black text-brand-primary uppercase tracking-tight mb-2">{step.title}</h4>
                  <p className="text-xs font-semibold text-brand-muted leading-relaxed uppercase tracking-wider">{step.description}</p>
                </div>
              </div>
            ))}
          </div>
          
          <div className="p-6 bg-brand-primary text-white flex items-center justify-between">
            <div className="flex items-center gap-3">
              <Sparkles className="w-4 h-4 text-brand-accent" />
              <p className="text-[10px] font-black uppercase tracking-widest">Powered by DCM Intelligence Core</p>
            </div>
            <button 
              onClick={onClose}
              className="px-6 py-2 bg-brand-accent text-white rounded-lg font-black text-[10px] uppercase tracking-widest hover:scale-105 transition-all shadow-lg shadow-brand-accent/20"
            >
              Got it
            </button>
          </div>
        </motion.div>
      </motion.div>
    </AnimatePresence>
  );
}
