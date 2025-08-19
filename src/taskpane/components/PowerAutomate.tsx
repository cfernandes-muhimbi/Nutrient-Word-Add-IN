import * as React from "react";

/* global Office, console, btoa */

interface PowerAutomateProps {
  isVisible: boolean;
  onSessionUrlReceived?: (url: string) => void;
}

interface PowerAutomateConfig {
  enabled: boolean;
  webhookUrl: string;
}

const PowerAutomate: React.FC<PowerAutomateProps> = ({ isVisible, onSessionUrlReceived }) => {
  const [isProcessing, setIsProcessing] = React.useState<boolean>(false);
  const [message, setMessage] = React.useState<string>("");
  const [messageType, setMessageType] = React.useState<string>("info");
  const [powerAutomateConfig, setPowerAutomateConfig] = React.useState<PowerAutomateConfig>({
    enabled: false,
    webhookUrl: ""
  });

  const showMessage = (text: string, type: string = "info") => {
    setMessage(text);
    setMessageType(type);
    setTimeout(() => setMessage(""), 4000);
  };

  const getDocumentAsBase64 = (): Promise<string> => {
    return new Promise((resolve, reject) => {
      const timeoutId = setTimeout(() => {
        reject(new Error("Export timed out after 5 minutes"));
      }, 300000);

      Office.context.document.getFileAsync(Office.FileType.Pdf, function (result) {
        clearTimeout(timeoutId);
        
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const myFile = result.value;
          console.log("Getting whole PDF file without slicing");
          
          myFile.getSliceAsync(0, function (sliceResult: any) {
            if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
              const slice = sliceResult.value;
              myFile.closeAsync();
              
              console.log("Whole file data type:", typeof slice.data);
              console.log("Whole file data length:", slice.data?.length || 0);
              
              if (typeof slice.data === 'string') {
                console.log("PDF data is already base64 string");
                resolve(slice.data);
              } else {
                console.log("Converting whole PDF binary data to base64");
                
                let bytes: Uint8Array;
                if (slice.data instanceof ArrayBuffer) {
                  bytes = new Uint8Array(slice.data);
                } else if (Array.isArray(slice.data)) {
                  bytes = new Uint8Array(slice.data);
                } else {
                  bytes = slice.data;
                }
                
                console.log("Processing", bytes.length, "bytes");
                
                try {
                  if (typeof Buffer !== 'undefined') {
                    const base64Data = Buffer.from(bytes).toString('base64');
                    console.log("Converted using Buffer, length:", base64Data.length);
                    resolve(base64Data);
                  } else {
                    let safeString = '';
                    for (let i = 0; i < bytes.length; i++) {
                      const byte = bytes[i] & 0xFF;
                      safeString += String.fromCharCode(byte);
                    }
                    const base64Data = btoa(safeString);
                    console.log("Converted using btoa, length:", base64Data.length);
                    resolve(base64Data);
                  }
                } catch (error) {
                  console.error("Failed to convert whole file to base64:", error);
                  reject(new Error("Failed to convert PDF to base64"));
                }
              }
            } else {
              myFile.closeAsync();
              console.error("Failed to get whole file:", sliceResult.error);
              reject(new Error(`Failed to get PDF data: ${sliceResult.error.message}`));
            }
          });
        } else {
          console.error("getFileAsync failed:", result.error);
          reject(new Error(`Failed to get document: ${result.error.message}`));
        }
      });
    });
  };

  const sendToPowerAutomate = async () => {
    if (!powerAutomateConfig.enabled || !powerAutomateConfig.webhookUrl) {
      showMessage("Please configure the Power Automate webhook URL first.", "error");
      return;
    }

    setIsProcessing(true);

    try {
      showMessage("Sending document to Power Automate...", "info");
      
      const docxBase64 = await getDocumentAsBase64();
      console.log("Document content retrieved as Clavin:", docxBase64);

      const payload = {
        fileName: "document.pdf",
        contentType: "application/pdf",
        fileContent: docxBase64,
        timestamp: new Date().toISOString(),
      };

      const response = await fetch(powerAutomateConfig.webhookUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      console.log("Power Automate Response Status:", response.status);
      console.log("Power Automate Response Headers:", response.headers);

      if (response.ok) {
        const responseBody = await response.text();
        console.log("Power Automate Response Body:", responseBody);
        
        try {
          const jsonResponse = JSON.parse(responseBody);
          console.log("Power Automate Response JSON:", jsonResponse);
          
          // Check if the response contains a session URL
          if (jsonResponse.sessionUrl && onSessionUrlReceived) {
            console.log("Session URL received from Power Automate:", jsonResponse.sessionUrl);
            onSessionUrlReceived(jsonResponse.sessionUrl);
          }
        } catch (e) {
          console.log("Response is not JSON, treating as text");
          // If response is plain text and looks like a URL, use it as session URL
          if (responseBody.includes('http') && onSessionUrlReceived) {
            console.log("Session URL received as text from Power Automate:", responseBody);
            onSessionUrlReceived(responseBody.trim());
          }
        }
        
        showMessage("Document sent to Power Automate successfully!", "success");
      } else {
        const errorBody = await response.text();
        console.log("Power Automate Error Response:", errorBody);
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
    } catch (error) {
      console.error("Error sending to Power Automate:", error);
      showMessage("Failed to send document to Power Automate. Check console for details.", "error");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleWebhookUrlChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const newUrl = event.target.value;
    setPowerAutomateConfig(prev => ({
      ...prev,
      webhookUrl: newUrl,
      enabled: newUrl.trim() !== ""
    }));
  };

  if (!isVisible) {
    return null;
  }

  return (
    <div 
      style={{
        padding: "20px",
        backgroundColor: "white",
        borderRadius: "8px",
        margin: "20px 0",
        boxShadow: "0 2px 4px rgba(0,0,0,0.1)",
        textAlign: "center"
      }}
    >
      <h3 
        style={{
          marginBottom: "15px",
          color: "#323130",
          fontSize: "18px",
          fontWeight: "600"
        }}
      >
        Power Automate Trigger
      </h3>

      <p style={{ marginBottom: "20px", fontSize: "14px", color: "#605e5c" }}>
        Send the current document to a Power Automate workflow
      </p>

      <div style={{ marginBottom: "20px" }}>
        <label 
          style={{
            display: "block",
            marginBottom: "8px",
            fontSize: "14px",
            color: "#323130",
            textAlign: "left"
          }}
        >
          Webhook URL:
        </label>
        <input
          type="url"
          value={powerAutomateConfig.webhookUrl}
          onChange={handleWebhookUrlChange}
          placeholder="https://prod-xx.eastus.logic.azure.com:443/..."
          style={{
            width: "100%",
            padding: "8px 12px",
            border: "1px solid #d2d0ce",
            borderRadius: "4px",
            fontSize: "14px",
            boxSizing: "border-box"
          }}
        />
      </div>
      
      <button
        onClick={sendToPowerAutomate}
        disabled={isProcessing || !powerAutomateConfig.enabled}
        style={{
          backgroundColor: isProcessing || !powerAutomateConfig.enabled ? "#ccc" : "#0078d4",
          color: "white",
          border: "none",
          padding: "15px 30px",
          borderRadius: "6px",
          cursor: isProcessing || !powerAutomateConfig.enabled ? "not-allowed" : "pointer",
          fontSize: "16px",
          fontWeight: "600",
          minWidth: "200px"
        }}
      >
        {isProcessing ? "Sending..." : "Send to Power Automate"}
      </button>

      {message && (
        <div 
          style={{
            marginTop: "15px",
            padding: "10px",
            backgroundColor: messageType === "error" ? "#fde7e9" : messageType === "success" ? "#dff6dd" : "#f3f2f1",
            borderRadius: "4px",
            border: `1px solid ${messageType === "error" ? "#a80000" : messageType === "success" ? "#107c10" : "#edebe9"}`
          }}
        >
          <p 
            style={{
              margin: "0",
              color: messageType === "error" ? "#a80000" : messageType === "success" ? "#107c10" : "#323130",
              fontSize: "13px"
            }}
          >
            {message}
          </p>
        </div>
      )}
    </div>
  );
};

export default PowerAutomate;