import * as React from "react";

interface ViewerProps {
  isVisible: boolean;
  documentUrl?: string;
  sessionUrl?: string;
  onDocumentLoaded?: () => void;
}

const Viewer: React.FC<ViewerProps> = ({ isVisible, documentUrl, sessionUrl, onDocumentLoaded }) => {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const viewerInstanceRef = React.useRef<any>(null);

  React.useEffect(() => {
    const container = containerRef.current;
    
    if (!isVisible || !container) {
      // Return cleanup function even when not loading
      return () => {
        if (viewerInstanceRef.current) {
          try {
            // Use dynamic import for cleanup as well
            import("@nutrient-sdk/viewer").then(({ default: NutrientViewer }) => {
              NutrientViewer.unload(container);
              viewerInstanceRef.current = null;
            }).catch(error => {
              console.warn('Error during viewer cleanup:', error);
            });
          } catch (error) {
            console.warn('Error during viewer cleanup:', error);
          }
        }
      };
    }

    const loadViewer = async () => {
      try {
        // Dynamic import of the Nutrient SDK
        const NutrientViewer = (await import("@nutrient-sdk/viewer")).default;

        // Clean up any existing instance
        if (viewerInstanceRef.current) {
          try {
            await NutrientViewer.unload(container);
          } catch (error) {
            console.warn('Error unloading previous viewer instance:', error);
          }
        }

        // Load the viewer with a document or session
        const loadConfig: any = {
          container,
          baseUrl: "https://cdn.cloud.pspdfkit.com/pspdfkit-web@1.6.0/",
          sessionUrl:sessionUrl
          //licenseKey: "YOUR_LICENSE_KEY_HERE" // Replace with actual license key
        };

        if (sessionUrl) {
          // Use session URL if provided
          loadConfig.session = sessionUrl;
          console.log('Loading viewer with session URL:', sessionUrl);
        } else {
          // Fallback to document URL
          const documentToLoad = documentUrl || "https://pspdfkitstorage12.blob.core.windows.net/$web/document (1).pdf";
          loadConfig.document = documentToLoad;
          console.log('Loading viewer with document URL:', documentToLoad);
        }
        
        viewerInstanceRef.current = await NutrientViewer.load(loadConfig);

        console.log('Nutrient Viewer loaded successfully');
        
        // Notify parent component that document is loaded
        if (onDocumentLoaded) {
          onDocumentLoaded();
        }
      } catch (error) {
        console.error('Error loading Nutrient Viewer:', error);
      }
    };

    loadViewer();

    // Cleanup function
    return () => {
      if (viewerInstanceRef.current) {
        try {
          // Use dynamic import for cleanup as well
          import("@nutrient-sdk/viewer").then(({ default: NutrientViewer }) => {
            NutrientViewer.unload(container);
            viewerInstanceRef.current = null;
          }).catch(error => {
            console.warn('Error during viewer cleanup:', error);
          });
        } catch (error) {
          console.warn('Error during viewer cleanup:', error);
        }
      }
    };
  }, [isVisible, documentUrl, sessionUrl]);

  if (!isVisible) {
    return null;
  }

  return (
    <div
      style={{
        width: "100%",
        height: "100vh"
      }}
    >
      <div
        ref={containerRef}
        style={{
          height: "100%",
          width: "100%"
        }}
      />
    </div>
  );
};

export default Viewer;