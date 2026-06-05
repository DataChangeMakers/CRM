import * as React from "react";
import { ErrorInfo, ReactNode } from "react";
import { AlertCircle, RefreshCw } from "lucide-react";

interface Props {
  children: ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
}

export class ErrorBoundary extends React.Component<Props, State> {
  state: State = {
    hasError: false,
    error: null
  };

  public static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error };
  }

  public componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error("Uncaught error:", error, errorInfo);
  }

  public render() {
    if (this.state.hasError) {
      let displayMessage = "Something went wrong.";
      try {
        const errorJson = JSON.parse(this.state.error?.message || "{}");
        if (errorJson.error) {
          displayMessage = `Security/Database Error: ${errorJson.error}`;
        }
      } catch {
        displayMessage = this.state.error?.message || displayMessage;
      }

      return (
        <div className="h-screen w-screen flex flex-col items-center justify-center bg-[#E4E3E0] p-6 text-center">
          <div className="max-w-md bg-white p-12 border border-[#141414] shadow-[8px_8px_0px_0px_rgba(20,20,20,1)]">
            <AlertCircle className="w-16 h-16 mx-auto mb-6 text-red-500" />
            <h1 className="text-2xl font-bold tracking-tighter mb-4">Application Error</h1>
            <p className="text-[#141414]/60 mb-8 font-mono text-sm break-words">
              {displayMessage}
            </p>
            <button 
              onClick={() => window.location.reload()}
              className="w-full bg-[#141414] text-white py-4 px-8 font-bold hover:bg-[#141414]/90 transition-all flex items-center justify-center gap-2"
            >
              <RefreshCw className="w-4 h-4" /> Reload Application
            </button>
          </div>
        </div>
      );
    }

    // @ts-ignore
    return this.props.children;
  }
}
