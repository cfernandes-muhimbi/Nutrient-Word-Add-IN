import * as React from "react";
import Viewer from "./Viewer";
import PowerAutomate from "./PowerAutomate";

const Logo: React.FC = () => {
  const [showViewer, setShowViewer] = React.useState<boolean>(false);
  const [showPowerAutomate, setShowPowerAutomate] = React.useState<boolean>(false);
  const [sessionUrl, setSessionUrl] = React.useState<string>("");
  const [isDocumentLoaded, setIsDocumentLoaded] = React.useState<boolean>(false);

  const toggleViewer = () => {
    setShowViewer(!showViewer);
  };

  const togglePowerAutomate = () => {
    setShowPowerAutomate(!showPowerAutomate);
  };

  const handlePowerAutomateSessionUrl = (url: string) => {
    setSessionUrl(url);
    // Automatically show the viewer when a session URL is received from Power Automate
    setShowViewer(true);
  };

  const handleDocumentLoaded = () => {
    setIsDocumentLoaded(true);
  };

  return (
    <div style={{ 
      textAlign: "center", 
      padding: "30px 20px", 
      backgroundColor: "#EFEBE7",
      minHeight: "100vh",
      fontFamily: "Segoe UI, Tahoma, Geneva, Verdana, sans-serif"
    }}>
      {!isDocumentLoaded && (
        <>
          <h1 style={{
            color: "#1A1414",
            fontSize: "24px",
            fontWeight: "600",
            marginBottom: "30px",
            marginTop: "0"
          }}>
            Nutrient DWS Viewer Add-In
          </h1>
          <img 
            src="../../assets/Nutrient.png" 
            alt="Nutrient Logo" 
            style={{ maxWidth: "200px", height: "auto", marginBottom: "30px" }}
          />
        </>
      )}
      
      <div style={{ display: "flex", gap: "10px", justifyContent: "center", marginBottom: "20px", flexWrap: "wrap" }}>
        <button
          onClick={toggleViewer}
          style={{
            backgroundColor: "#1A1414",
            color: "#EFEBE7",
            border: "none",
            padding: "12px 24px",
            borderRadius: "4px",
            cursor: "pointer",
            fontSize: "14px",
            fontWeight: "600"
          }}
        >
          {showViewer ? "Hide" : "Show"} Document Viewer
        </button>
        
        <button
          onClick={togglePowerAutomate}
          style={{
            backgroundColor: "#0078d4",
            color: "white",
            border: "none",
            padding: "12px 24px",
            borderRadius: "4px",
            cursor: "pointer",
            fontSize: "14px",
            fontWeight: "600"
          }}
        >
          {showPowerAutomate ? "Hide" : "Show"} Power Automate
        </button>
      </div>
      
      <Viewer 
        isVisible={showViewer}
        sessionUrl={sessionUrl}
        onDocumentLoaded={handleDocumentLoaded}
      />
      
      <PowerAutomate 
        isVisible={showPowerAutomate}
        onSessionUrlReceived={handlePowerAutomateSessionUrl}
      />
    </div>
  );
};

export default Logo;