import { GoogleGenAI, Type } from "@google/genai";

export interface NonProfitResearch {
  name: string;
  revenue: string;
  products_services: string;
  staff_members: string;
  mission: string;
  website: string;
  linkedin_activity: string;
  staff_linkedin_summary: string;
  linkedin_url?: string;
  linkedin_overview?: string;
  ein?: string;
  propublica_grants?: string;
  charity_navigator_rating?: string;
}

// Cache for AI instances
const aiInstances: Record<string, GoogleGenAI> = {};

const MODEL_MAP: Record<string, string> = {
  'gemini-3-flash-preview': 'gemini-3-flash-preview',
  'gemini-3.1-pro-preview': 'gemini-3.1-pro-preview',
  'gemini-flash-latest': 'gemini-flash-latest',
  'gemini-pro-latest': 'gemini-pro-latest',
  'gemini-3.1-flash-lite-preview': 'gemini-3.1-flash-lite-preview',
  'gemini-1.5-flash': 'gemini-3-flash-preview',
  'gemini-1.5-pro': 'gemini-3-flash-preview',
  'gemini-1.5-flash-8b': 'gemini-3-flash-preview' 
};

function getAI(apiKey?: string) {
  const keyToUse = apiKey || process.env.GEMINI_API_KEY;
  if (!keyToUse) {
    throw new Error("Gemini API key is required but not configured. Please provide one in AI Config.");
  }

  if (!aiInstances[keyToUse]) {
    aiInstances[keyToUse] = new GoogleGenAI({ apiKey: keyToUse });
  }
  return aiInstances[keyToUse];
}

function resolveModel(model?: string): string {
  const base = model || "gemini-3-flash-preview";
  return MODEL_MAP[base] || base;
}

export async function researchNonProfit(
  orgName: string, 
  websiteUrl?: string, 
  config?: { model?: string; sources?: string[]; apiKey?: string }
): Promise<NonProfitResearch> {
  const modelName = resolveModel(config?.model);
  const sourcesText = config?.sources && config.sources.length > 0
    ? ` Prioritize these specific websites and sources for your search: ${config.sources.join(", ")}.`
    : "";

  try {
    const ai = getAI(config?.apiKey);
    
    const prompt = `CRITICAL IDENTITY CHECK: You are researching the non-profit organization named "${orgName}"${websiteUrl ? ` (Official Website: ${websiteUrl})` : ""}.${sourcesText}
            
            Find information about their:
            - Official EIN (Tax ID number for US non-profits): Verify it strictly belongs to ${orgName}.
            - Annual revenue (latest available)
            - Products or services they provide
            - Key staff members or leadership
            - Their mission statement
            - Official website URL
            - Recent LinkedIn activity (style, content, engagement)
            - A summary list of key staff based on their available LinkedIn info
            - Grant history and status (e.g. from ProPublica)
            - Rating from Charity Navigator
            
            STRICT DATA SEPARATION: DO NOT return financial or legal data for unrelated entities even if they have similar names or EINs. If specific data for "${orgName}" is not found in a source, explicitly state "No public record found for this organization".
            
            Return the result strictly as a valid JSON object. ALL VALUES MUST BE STRINGS.`;

    const response = await ai.models.generateContent({
      model: modelName,
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            name: { type: Type.STRING },
            revenue: { type: Type.STRING },
            products_services: { type: Type.STRING },
            staff_members: { type: Type.STRING },
            mission: { type: Type.STRING },
            website: { type: Type.STRING },
            linkedin_activity: { type: Type.STRING },
            staff_linkedin_summary: { type: Type.STRING },
            ein: { type: Type.STRING },
            propublica_grants: { type: Type.STRING },
            charity_navigator_rating: { type: Type.STRING }
          },
          required: ["name", "revenue", "products_services", "staff_members", "mission", "website", "linkedin_activity", "staff_linkedin_summary", "ein", "propublica_grants", "charity_navigator_rating"]
        }
      }
    });

    const baseResult: NonProfitResearch = JSON.parse(response.text || "{}");

    // Complement with ProPublica if EIN found
    if (baseResult.ein && typeof baseResult.ein === 'string') {
      try {
        const cleanEin = baseResult.ein.replace(/[^0-9]/g, '');
        if (cleanEin) {
          const ppResponse = await fetch(`/api/propublica/${cleanEin}`);
          if (ppResponse.ok) {
            const ppData = await ppResponse.json();
            console.log("ProPublica data fetched:", ppData);
          }
        }
      } catch (e) {
        console.error("ProPublica fetch failed", e);
      }
    }

    return baseResult;
  } catch (error: any) {
    console.error("Gemini researchNonProfit failed:", error);
    throw error;
  }
}

export async function researchLinkedIn(url: string, orgName: string, config?: { model?: string; apiKey?: string }): Promise<{ overview: string, topPeople: string, topPosts: string }> {
  try {
    const modelName = resolveModel(config?.model);
    const ai = getAI(config?.apiKey);
    
    const prompt = `Using focus on the LinkedIn profile at "${url}" for "${orgName}", provide:
          1. Overview
          2. Top People
          3. Top Posts
          Return strictly as JSON with keys: overview, topPeople, topPosts.`;

    const response = await ai.models.generateContent({
      model: modelName,
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            overview: { type: Type.STRING },
            topPeople: { type: Type.STRING },
            topPosts: { type: Type.STRING }
          },
          required: ["overview", "topPeople", "topPosts"]
        }
      }
    });

    return JSON.parse(response.text || "{}");
  } catch (error: any) {
    console.error("Gemini researchLinkedIn failed:", error);
    throw error;
  }
}

export interface ScoringDetail {
  value: number | boolean;
  comment: string;
}

export interface FullAssessmentResult {
  verification: {
    reputable: ScoringDetail;
    fitsMission: ScoringDetail;
    isNonProfit: ScoringDetail;
    available: ScoringDetail;
  };
  validation: {
    dataQuality: ScoringDetail;
    problemStatement: ScoringDetail;
    missionAlignment: ScoringDetail;
    partnershipReason: ScoringDetail;
    fundsAvailable: ScoringDetail;
  };
  validationChecks: {
    diverseStaff: ScoringDetail;
    diverseDemographic: ScoringDetail;
    inclusiveMarketing: ScoringDetail;
  };
}

export async function generateFullAssessment(
  name: string,
  mission: string,
  services: string,
  context: string = "",
  briefSummary: string = "",
  config?: { model?: string; apiKey?: string; fullText?: string }
): Promise<FullAssessmentResult> {
  const modelName = resolveModel(config?.model);
  try {
    const ai = getAI(config?.apiKey);
    
    const prompt = `Act as a senior project evaluator for "Changemaker Systems". 
          Analyze the provided information and potentially a full project brief document to fill out the assessment rubrics.
          
          Partner Info:
          Name: ${name}
          Mission: ${mission}
          Services: ${services}
          Context: ${context}
          Short Brief Summary: ${briefSummary}
          
          ${config?.fullText ? `FULL PROJECT BRIEF DOCUMENT TEXT:\n${config.fullText}` : ""}
          
          Return strictly JSON. Each sub-key must have a "value" (boolean for verification/checks, 1-4 for validation) and a "comment" detailing why that score/value was given based on the evidence in the text.
          `;

    const response = await ai.models.generateContent({
      model: modelName,
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            verification: {
              type: Type.OBJECT,
              properties: {
                reputable: { type: Type.OBJECT, properties: { value: { type: Type.BOOLEAN }, comment: { type: Type.STRING } }, required: ["value", "comment"] },
                fitsMission: { type: Type.OBJECT, properties: { value: { type: Type.BOOLEAN }, comment: { type: Type.STRING } }, required: ["value", "comment"] },
                isNonProfit: { type: Type.OBJECT, properties: { value: { type: Type.BOOLEAN }, comment: { type: Type.STRING } }, required: ["value", "comment"] },
                available: { type: Type.OBJECT, properties: { value: { type: Type.BOOLEAN }, comment: { type: Type.STRING } }, required: ["value", "comment"] }
              },
              required: ["reputable", "fitsMission", "isNonProfit", "available"]
            },
            validation: {
              type: Type.OBJECT,
              properties: {
                dataQuality: { type: Type.OBJECT, properties: { value: { type: Type.NUMBER }, comment: { type: Type.STRING } }, required: ["value", "comment"] },
                problemStatement: { type: Type.OBJECT, properties: { value: { type: Type.NUMBER }, comment: { type: Type.STRING } }, required: ["value", "comment"] },
                missionAlignment: { type: Type.OBJECT, properties: { value: { type: Type.NUMBER }, comment: { type: Type.STRING } }, required: ["value", "comment"] },
                partnershipReason: { type: Type.OBJECT, properties: { value: { type: Type.NUMBER }, comment: { type: Type.STRING } }, required: ["value", "comment"] },
                fundsAvailable: { type: Type.OBJECT, properties: { value: { type: Type.NUMBER }, comment: { type: Type.STRING } }, required: ["value", "comment"] }
              },
              required: ["dataQuality", "problemStatement", "missionAlignment", "partnershipReason", "fundsAvailable"]
            },
            validationChecks: {
              type: Type.OBJECT,
              properties: {
                diverseStaff: { type: Type.OBJECT, properties: { value: { type: Type.BOOLEAN }, comment: { type: Type.STRING } }, required: ["value", "comment"] },
                diverseDemographic: { type: Type.OBJECT, properties: { value: { type: Type.BOOLEAN }, comment: { type: Type.STRING } }, required: ["value", "comment"] },
                inclusiveMarketing: { type: Type.OBJECT, properties: { value: { type: Type.BOOLEAN }, comment: { type: Type.STRING } }, required: ["value", "comment"] }
              },
              required: ["diverseStaff", "diverseDemographic", "inclusiveMarketing"]
            }
          },
          required: ["verification", "validation", "validationChecks"]
        }
      }
    });

    return JSON.parse(response.text || "{}");
  } catch (error: any) {
    console.error("Gemini generateFullAssessment failed:", error);
    throw error;
  }
}

export async function summarizeBrief(orgName: string, rawContent: string, config?: { model?: string; apiKey?: string }): Promise<string> {
  try {
    const modelName = resolveModel(config?.model);
    const ai = getAI(config?.apiKey);
    const response = await ai.models.generateContent({
      model: modelName,
      contents: `Summarize this project brief for "${orgName}": ${rawContent}. Max 200 words.`
    });
    return response.text || "";
  } catch (error: any) {
    console.error("Gemini summarizeBrief failed:", error);
    throw error;
  }
}

export interface DataQualityAssessment {
  score: number;
  verdict: string;
  strengths: string[];
  weaknesses: string[];
  recommendations: string[];
}

export async function assessDataQuality(orgName: string, dataSample: string, config?: { model?: string; apiKey?: string }): Promise<DataQualityAssessment> {
  try {
    const modelName = resolveModel(config?.model);
    const ai = getAI(config?.apiKey);
    
    const prompt = `Assess data from "${orgName}": ${dataSample}. 
          Return strictly JSON with keys: score (1-10), verdict, strengths (array), weaknesses (array), recommendations (array).`;

    const response = await ai.models.generateContent({
      model: modelName,
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            score: { type: Type.NUMBER },
            verdict: { type: Type.STRING },
            strengths: { type: Type.ARRAY, items: { type: Type.STRING } },
            weaknesses: { type: Type.ARRAY, items: { type: Type.STRING } },
            recommendations: { type: Type.ARRAY, items: { type: Type.STRING } }
          },
          required: ["score", "verdict", "strengths", "weaknesses", "recommendations"]
        }
      }
    });

    return JSON.parse(response.text || "{}");
  } catch (error: any) {
    console.error("Gemini assessDataQuality failed:", error);
    throw error;
  }
}

export interface LeadAction {
  leadId: string;
  organisation: string;
  action: string;
  priority: 'high' | 'medium' | 'low';
  context: string;
  leadOwner: string;
}

export async function suggestFollowUpEmail(lead: any, userProfile?: any, config?: { model?: string; apiKey?: string }): Promise<string> {
  try {
    const modelName = resolveModel(config?.model);
    const ai = getAI(config?.apiKey);
    const response = await ai.models.generateContent({
      model: modelName,
      contents: `Draft a follow-up email for "${lead.organisation}". Context: ${JSON.stringify(lead)}. User Profile: ${JSON.stringify(userProfile)}.`
    });
    return response.text || "";
  } catch (error: any) {
    console.error("Gemini suggestFollowUpEmail failed:", error);
    throw error;
  }
}

export async function extractActionsFromLeads(leads: any[], config?: { model?: string; apiKey?: string }): Promise<LeadAction[]> {
  try {
    const modelName = resolveModel(config?.model);
    const ai = getAI(config?.apiKey);
    
    const prompt = `Extract actionable CRM insights from these leads: ${JSON.stringify(leads)}. 
          Return strictly a JSON array of objects with keys: leadId, organisation, action, priority (high/medium/low), context, leadOwner.`;

    const response = await ai.models.generateContent({
      model: modelName,
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              leadId: { type: Type.STRING },
              organisation: { type: Type.STRING },
              action: { type: Type.STRING },
              priority: { type: Type.STRING },
              context: { type: Type.STRING },
              leadOwner: { type: Type.STRING }
            },
            required: ["leadId", "organisation", "action", "priority", "context", "leadOwner"]
          }
        },
      }
    });

    return JSON.parse(response.text || "[]");
  } catch (error: any) {
    console.error("Gemini extractActionsFromLeads failed:", error);
    throw error;
  }
}
