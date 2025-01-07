import * as React from "react";
import * as ReactWebChat from "botframework-webchat";
import { Dialog } from "@fluentui/react/lib/Dialog";
import { IconButton } from "@fluentui/react/lib/Button";
import { Spinner } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Dispatch } from "redux";
import { useRef, useEffect, useState } from "react";
import { IChatbotProps } from "./IChatBotProps";
import MSALWrapper from "./MSALWrapper";
import styles from "./PVAChatbotDialog.module.scss";

export const PVAChatbotDialog: React.FC<IChatbotProps> = (props) => {
  const [error, setError] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const webChatRef = useRef<HTMLDivElement>(null);

  const modalProps = React.useMemo(() => ({
    isBlocking: false,
    className: styles.dialogRoot,
  }), []);

  const parseBotUrl = (url: string) => {
    try {
      const parsedUrl = new URL(url);
      const apiVersionParam = parsedUrl.searchParams.get("api-version");

      if (!apiVersionParam) {
        throw new Error("Missing api-version parameter");
      }

      return {
        isValid: true,
        environmentEndPoint: parsedUrl.origin,
        apiVersion: apiVersionParam,
        regionalChannelSettingsURL: `${parsedUrl.origin}/powervirtualagents/regionalchannelsettings?api-version=${apiVersionParam}`,
      };
    } catch (err) {
      console.error("Error parsing bot URL:", err);
      setError("Invalid bot URL configuration");
      return {
        isValid: false,
        environmentEndPoint: "",
        apiVersion: "",
        regionalChannelSettingsURL: "",
      };
    }
  };

  const handleLayerDidMount = async () => {
    const botURL = props.botURL?.trim();
    if (!botURL) {
      setError("Bot URL is not configured");
      setIsLoading(false);
      return;
    }

    const urlInfo = parseBotUrl(botURL);
    if (!urlInfo.isValid || !urlInfo.regionalChannelSettingsURL) {
      setError("Invalid bot URL configuration");
      setIsLoading(false);
      return;
    }

    try {
      const styleOptions = {
        hideUploadButton: true,
        backgroundColor: "transparent",
        botAvatarBackgroundColor: "var(--themePrimary)",
        botAvatarInitials: props.botAvatarInitials || "BOT",
        userAvatarBackgroundColor: "var(--neutral-tertiary)",
        userAvatarInitials: props.userFriendlyName?.split(" ").map((n) => n[0]).join("") || "U",
        bubbleBackground: "var(--neutral-lighter)",
        bubbleFromUserBackground: "var(--theme-lighter-alt)",
        bubbleBorderRadius: 8,
        bubbleFromUserBorderRadius: 8,
        bubblePadding: 12,
        rootHeight: "100%",
        rootWidth: "100%",
        sendBoxBackground: "var(--white)",
        sendBoxBorderTop: "1px solid var(--neutralLight)",
        sendBoxTextWrap: true,
        transitionDuration: "0.2s",
      };

      const MSALWrapperInstance = new MSALWrapper(props.clientID, props.authority);

      let responseToken = await MSALWrapperInstance.handleLoggedInUser([props.customScope], props.userEmail);
      if (!responseToken) {
        responseToken = await MSALWrapperInstance.acquireAccessToken([props.customScope], props.userEmail);
      }
      const token = responseToken?.accessToken;

      if (!token) {
        throw new Error("Failed to acquire access token");
      }

      const regionalResponse = await fetch(urlInfo.regionalChannelSettingsURL);
      if (!regionalResponse.ok) {
        throw new Error(`Failed to fetch regional settings: ${regionalResponse.status}`);
      }

      const regionalData = await regionalResponse.json();
      const regionalChannelURL = regionalData.channelUrlsById?.directline;
      if (!regionalChannelURL) {
        throw new Error("DirectLine URL not found");
      }

      const response = await fetch(botURL);
      if (!response.ok) {
        throw new Error(`Failed to fetch DirectLine token: ${response.status}`);
      }

      const conversationInfo = await response.json();
      const directline = ReactWebChat.createDirectLine({
        token: conversationInfo.token,
        domain: `${regionalChannelURL}v3/directline`,
      });

      const store = ReactWebChat.createStore({}, ({ dispatch }: { dispatch: Dispatch }) =>
        (next: any) => (action: any) => {
          if (props.greet && action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
            dispatch({
              meta: { method: "keyboard" },
              payload: {
                activity: {
                  channelData: { postBack: true },
                  name: "startConversation",
                  type: "event",
                },
              },
              type: "DIRECT_LINE/POST_ACTIVITY",
            });
          }
          return next(action);
        }
      );

      if (webChatRef.current) {
        await ReactWebChat.renderWebChat(
          {
            directLine: directline,
            store: store,
            styleOptions,
            userID: props.userEmail,
          },
          webChatRef.current
        );
      }
    } catch (err) {
      setError((err as Error).message);
      console.error("Chat setup error:", err);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    if (!props.isOpen) return;
  
    setIsLoading(true);
    setError(null);
  
    handleLayerDidMount()
      .then(() => {
        // Successful initialization (if needed, handle success here)
      })
      .catch((error) => {
        console.error("Error initializing chatbot:", error);
        setError("Failed to initialize chatbot."); // Optionally set error state
      });
  
    return () => {
      // Cleanup: Clear the Web Chat container
      if (webChatRef.current) {
        webChatRef.current.innerHTML = ""; // Clear the container
      }
    };
  }, [props.isOpen]);
  
  return (
    <Dialog
      hidden={!props.isOpen}
      onDismiss={props.onDismiss}
      modalProps={modalProps}
    >
      <div className={styles.msDialogHeader}>
        <span className={styles.title}>{props.botName}</span>
        <IconButton
          className={styles.closeButton}
          iconProps={{ iconName: "Cancel" }}
          ariaLabel="Close dialog"
          onClick={props.onDismiss}
        />
      </div>
      <div className={styles.dialogContent}>
        <div className={styles.contentContainer}>
          {error && (
            <div className={styles.errorContainer}>
              <MessageBar
                messageBarType={MessageBarType.error}
                onDismiss={() => setError(null)}
                dismissButtonAriaLabel="Close"
              >
                {error}
              </MessageBar>
            </div>
          )}

          <div className={styles.webChat} ref={webChatRef} role="main" />

          {isLoading && (
            <div className={styles.spinnerContainer}>
              <Spinner label="Loading chat..." labelPosition="bottom" size={3} />
            </div>
          )}
        </div>
      </div>
    </Dialog>
  );
};  