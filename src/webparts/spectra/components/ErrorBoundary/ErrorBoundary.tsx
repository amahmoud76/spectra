import * as React from "react";
import { captureAndLogError } from "../../services/errorLogService";

interface IErrorBoundaryProps {
  children: React.ReactNode;
}

interface IErrorBoundaryState {
  hasError: boolean;
  error: Error | undefined;
}

/**
 * Error boundary component to catch React errors and log them
 */
export class ErrorBoundary extends React.Component<
  IErrorBoundaryProps,
  IErrorBoundaryState
> {
  constructor(props: IErrorBoundaryProps) {
    super(props);
    this.state = {
      hasError: false,
      error: undefined,
    };
  }

  static getDerivedStateFromError(error: Error): IErrorBoundaryState {
    return {
      hasError: true,
      error,
    };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    // Log the error to SharePoint asynchronously
    captureAndLogError(error, {
      component: "Error Boundary",
      errorType: "React Error",
      additionalContext: {
        componentStack: errorInfo.componentStack,
      },
    }).catch(() => undefined);

    console.error("Error caught by boundary:", error, errorInfo);
  }

  render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <div
          style={{
            padding: "20px",
            margin: "10px",
            backgroundColor: "#f8d7da",
            border: "1px solid #f5c6cb",
            borderRadius: "4px",
            color: "#721c24",
          }}
        >
          <h2>Something went wrong</h2>
          <p>An error has been logged and the admin has been notified.</p>
          <p style={{ fontSize: "12px", marginTop: "10px" }}>
            Error: {this.state.error?.message}
          </p>
          <button
            onClick={() => window.location.reload()}
            style={{
              marginTop: "10px",
              padding: "8px 16px",
              backgroundColor: "#721c24",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: "pointer",
            }}
          >
            Reload Page
          </button>
        </div>
      );
    }

    return this.props.children;
  }
}
