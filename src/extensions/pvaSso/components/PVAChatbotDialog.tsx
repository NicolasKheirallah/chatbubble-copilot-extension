import * as React from "react";
import { useId } from '@fluentui/react-hooks';
import ReactWebChat, { createDirectLine, createStore } from 'botframework-webchat';
import { FluentThemeProvider } from 'botframework-webchat-fluent-theme';
import { Dialog, DialogType } from '@fluentui/react/lib/Dialog';
import { IconButton } from '@fluentui/react/lib/Button';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { Dispatch } from 'redux';
import { useRef, useEffect, useMemo, useState } from "react";
import { IChatbotProps } from "../types/IChatBotProps";
import MSALWrapper from "../services/MSALWrapper";
import styles from "../styles/PvaSsoApplicationCustomizer.module.scss";

// TypeScript declarations for Web Speech API
declare global {
  interface Window {
    SpeechRecognition: any;
    webkitSpeechRecognition: any;
    SpeechGrammarList: any;
    SpeechSynthesisUtterance: any;
  }
}

export const PVAChatbotDialog: React.FC<IChatbotProps> = (props) => {
  const dialogContentProps = {
    type: DialogType.normal,
    title: (
      <div className={styles.header} style={{ 
        backgroundColor: '#0078d4 !important', 
        background: '#0078d4 !important', 
        color: 'white !important',
        borderBottom: 'none !important',
        boxShadow: 'none !important'
      }}>
        <IconButton 
          iconProps={{ iconName: 'History' }}
          ariaLabel="Session History"
          onClick={() => setShowSessionHistory(!showSessionHistory)}
          className={styles.historyButton}
          styles={{
            root: {
              color: 'white !important',
              backgroundColor: 'transparent !important',
              marginRight: '8px',
              border: 'none !important',
              '&:hover': {
                backgroundColor: 'rgba(255, 255, 255, 0.1) !important',
                color: 'white !important'
              }
            },
            icon: { color: 'white !important', fontSize: '16px !important' },
            flexContainer: { color: 'white !important' }
          }}
        />
        <span className={styles.chatTitle}>{props.botName}</span>
        <IconButton 
          iconProps={{ iconName: 'Add' }}
          ariaLabel="New Chat"
          onClick={() => {
            const newSessionId = `session-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
            setSessionId(newSessionId);
            setCurrentSessionData(null);
            localStorage.removeItem(currentSessionKey);
            // Force WebChat to reinitialize by ending the current DirectLine connection
            if (directLine) {
              try {
                directLine.end();
              } catch (error) {
                console.warn('Error ending DirectLine connection for new chat:', error);
              }
            }
          }}
          className={styles.newChatButton}
          styles={{
            root: {
              color: 'white !important',
              backgroundColor: 'transparent !important',
              marginRight: '8px',
              border: 'none !important',
              '&:hover': {
                backgroundColor: 'rgba(255, 255, 255, 0.1) !important',
                color: 'white !important'
              }
            },
            icon: { color: 'white !important', fontSize: '16px !important' },
            flexContainer: { color: 'white !important' }
          }}
        />
        <IconButton 
          iconProps={{ iconName: 'Cancel' }}
          ariaLabel="Close chat"
          onClick={props.onDismiss}
          className={styles.chatCloseButton}
          styles={{
            root: {
              color: 'white !important',
              backgroundColor: 'transparent !important',
              border: 'none !important',
              boxShadow: 'none !important',
              '&:hover': {
                backgroundColor: 'rgba(255, 255, 255, 0.1) !important',
                color: 'white !important'
              },
              '&:focus': {
                backgroundColor: 'rgba(255, 255, 255, 0.1) !important',
                color: 'white !important',
                outline: '2px solid rgba(255, 255, 255, 0.5)'
              },
              '&:active': {
                backgroundColor: 'rgba(255, 255, 255, 0.2) !important',
                color: 'white !important'
              }
            },
            icon: {
              fontSize: '16px !important',
              color: 'white !important'
            },
            flexContainer: {
              color: 'white !important'
            }
          }}
        />
      </div>
    ),
    closeButtonAriaLabel: 'Close',
  };

  const labelId: string = useId('dialogLabel');
  const subTextId: string = useId('subTextLabel');

  const modalProps = React.useMemo(() => ({
    isBlocking: false
  }), [labelId, subTextId]);

  const botURL = props.botURL?.trim() || '';
  if (!botURL) {
    console.error("botURL is empty in PVAChatbotDialog. Check your props!");
  }
  const idx = botURL.indexOf('/powervirtualagents');
  if (idx === -1 && botURL !== '') {
    console.error("botURL doesn't contain '/powervirtualagents'. Check your config:", botURL);
  }

  const environmentEndPoint = idx > -1 ? botURL.slice(0, idx) : '';
  const queryIndex = botURL.indexOf('api-version');
  let apiVersion = "";
  if (queryIndex !== -1) {
    const versionPart = botURL.slice(queryIndex);
    const split = versionPart.split('=');
    apiVersion = split[1] || "";
  }

  const regionalChannelSettingsURL = environmentEndPoint && apiVersion
    ? `${environmentEndPoint}/powervirtualagents/regionalchannelsettings?api-version=${apiVersion}`
    : '';

  const loadingSpinnerRef = useRef<HTMLDivElement>(null);
  const msalWrapperRef = useRef<MSALWrapper | null>(null);
  const cleanupRef = useRef<(() => void) | null>(null);
  
  const [directLineToken, setDirectLineToken] = useState<string | null>(null);
  const [regionalChannelURL, setRegionalChannelURL] = useState<string | null>(null);
  const [sessionId, setSessionId] = useState<string>('');
  const [showSessionHistory, setShowSessionHistory] = useState<boolean>(false);
  const [sessionHistory, setSessionHistory] = useState<SessionData[]>([]);
  const [currentSessionData, setCurrentSessionData] = useState<SessionData | null>(null);
  
  interface SessionData {
    id: string;
    timestamp: number;
    title: string;
    messageCount: number;
    lastMessage?: string;
    activities: any[];
  }
  
  const sessionsStorageKey = 'webchat-sessions-history';
  const currentSessionKey = `webchat-current-session-${props.botURL.replace(/[^a-zA-Z0-9]/g, '-')}`;
  
  const saveCurrentSession = (activities: any[]) => {
    if (!sessionId || !activities.length) return;
    
    try {
      const sessionData: SessionData = {
        id: sessionId,
        timestamp: Date.now(),
        title: generateSessionTitle(activities),
        messageCount: activities.filter(a => a.type === 'message').length,
        lastMessage: getLastUserMessage(activities),
        activities
      };
      
      localStorage.setItem(currentSessionKey, JSON.stringify(sessionData));
      updateSessionHistory(sessionData);
      setCurrentSessionData(sessionData);
    } catch (error) {
      console.warn('Could not save session data:', error);
    }
  };
  
  const updateSessionHistory = (newSession: SessionData) => {
    try {
      const existing = localStorage.getItem(sessionsStorageKey);
      const sessions: SessionData[] = existing ? JSON.parse(existing) : [];
      
      const existingIndex = sessions.findIndex(s => s.id === newSession.id);
      if (existingIndex >= 0) {
        sessions[existingIndex] = newSession;
      } else {
        sessions.unshift(newSession);
      }
      
      const maxSessions = 20;
      const recentSessions = sessions.slice(0, maxSessions);
      
      localStorage.setItem(sessionsStorageKey, JSON.stringify(recentSessions));
      setSessionHistory(recentSessions);
    } catch (error) {
      console.warn('Could not update session history:', error);
    }
  };
  
  const loadSessionHistory = () => {
    try {
      const stored = localStorage.getItem(sessionsStorageKey);
      if (stored) {
        const sessions: SessionData[] = JSON.parse(stored);
        setSessionHistory(sessions.filter(s => Date.now() - s.timestamp < 7 * 24 * 60 * 60 * 1000));
      }
    } catch (error) {
      console.warn('Could not load session history:', error);
    }
  };
  
  const loadCurrentSession = () => {
    try {
      const stored = localStorage.getItem(currentSessionKey);
      if (stored) {
        const sessionData: SessionData = JSON.parse(stored);
        const dayAgo = 24 * 60 * 60 * 1000;
        if (Date.now() - sessionData.timestamp < dayAgo) {
          setCurrentSessionData(sessionData);
          return sessionData;
        }
        localStorage.removeItem(currentSessionKey);
      }
    } catch (error) {
      console.warn('Could not load current session:', error);
    }
    return null;
  };
  
  const generateSessionTitle = (activities: any[]): string => {
    const userMessages = activities.filter(a => a.type === 'message' && a.from?.role === 'user');
    if (userMessages.length > 0) {
      const firstMessage = userMessages[0].text || '';
      return firstMessage.length > 30 ? firstMessage.substring(0, 30) + '...' : firstMessage;
    }
    return `Chat ${new Date().toLocaleString()}`;
  };
  
  const getLastUserMessage = (activities: any[]): string => {
    const userMessages = activities.filter(a => a.type === 'message' && a.from?.role === 'user');
    return userMessages.length > 0 ? userMessages[userMessages.length - 1].text || '' : '';
  };
  
  
  useEffect(() => {
    loadSessionHistory();
    const stored = loadCurrentSession();
    if (stored?.id) {
      setSessionId(stored.id);
    } else {
      const newSessionId = `session-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
      setSessionId(newSessionId);
      setCurrentSessionData(null);
    }
  }, []);
  
  const webSpeechPonyfillFactory = useMemo(() => {
    if (!props.enableSpeechSynthesis) return undefined;
    
    return () => {
      if (typeof window !== 'undefined' && window.speechSynthesis) {
        return {
          SpeechGrammarList: window.SpeechGrammarList,
          SpeechRecognition: window.SpeechRecognition || window.webkitSpeechRecognition,
          speechSynthesis: window.speechSynthesis,
          SpeechSynthesisUtterance: window.SpeechSynthesisUtterance
        };
      }
      return {};
    };
  }, [props.enableSpeechSynthesis]);
  
  const attachmentMiddleware = useMemo(() => () => (next: any) => (renderProps: any) => {
    const { attachment } = renderProps;
    
    if (attachment.contentType?.startsWith('image/')) {
      return next({
        ...renderProps,
        alt: attachment.name || 'Image attachment',
        className: styles.imageAttachment
      });
    }
    
    if (attachment.contentType?.startsWith('video/')) {
      return next({
        ...renderProps,
        controls: true,
        className: styles.videoAttachment
      });
    }
    
    if (attachment.contentType?.startsWith('audio/')) {
      return next({
        ...renderProps,
        controls: true,
        className: styles.audioAttachment
      });
    }
    
    return next(renderProps);
  }, []);
  
  const directLine = useMemo(() => {
    if (!directLineToken || !regionalChannelURL) return null;
    
    try {
      return createDirectLine({
        token: directLineToken,
        domain: `${regionalChannelURL}v3/directline`
      });
    } catch (error) {
      console.error("DirectLine creation error:", error);
      return null;
    }
  }, [directLineToken, regionalChannelURL]);

  function getOAuthCardResourceUri(activity: any): string | undefined {
    const attachment = activity?.attachments?.[0];
    if (attachment?.contentType === 'application/vnd.microsoft.card.oauth' && attachment.content.tokenExchangeResource) {
      return attachment.content.tokenExchangeResource.uri;
    }
  }

  const handleLayerDidMount = async () => {
    if (!botURL || idx === -1 || !regionalChannelSettingsURL) {
      console.error("Invalid botURL or regionalChannelSettingsURL. Cannot set up chat.");
      return;
    }

    if (!msalWrapperRef.current) {
      msalWrapperRef.current = new MSALWrapper(props.clientID, props.authority);
    }
    const MSALWrapperInstance = msalWrapperRef.current;

    let responseToken = await MSALWrapperInstance.handleLoggedInUser([props.customScope], props.userEmail);
    if (!responseToken) {
      responseToken = await MSALWrapperInstance.acquireAccessToken([props.customScope], props.userEmail);
    }
    const token = responseToken?.accessToken || null;

    if (!token) {
      console.error("Failed to acquire access token.");
      return;
    }

    const regionalResponse = await fetch(regionalChannelSettingsURL);
    if (regionalResponse.ok) {
      const data = await regionalResponse.json();
      const channelURL = data.channelUrlsById?.directline;
      if (!channelURL) {
        console.error("DirectLine URL not found in regional channel settings.");
        return;
      }
      setRegionalChannelURL(channelURL);
    } else {
      console.error(`HTTP error fetching ${regionalChannelSettingsURL}: Status ${regionalResponse.status}`);
      return;
    }

    const response = await fetch(botURL);
    if (response.ok) {
      const conversationInfo = await response.json();
      if (conversationInfo && conversationInfo.token) {
        setDirectLineToken(conversationInfo.token);
      } else {
        console.error("Invalid conversation info received");
        return;
      }
    } else {
      console.error(`HTTP error fetching botURL: Status ${response.status}`);
      return;
    }

    if (!directLineToken) {
      console.error("DirectLine token not available");
      return;
    }
    
    if (loadingSpinnerRef.current) {
      loadingSpinnerRef.current.classList.add(styles.loadingSpinnerHidden);
    }

    cleanupRef.current = () => {
      if (directLine) {
        try {
          directLine.end();
        } catch (error) {
          console.warn('Error ending DirectLine connection:', error);
        }
      }
      
      if (loadingSpinnerRef.current) {
        loadingSpinnerRef.current.classList.remove(styles.loadingSpinnerHidden);
      }
    };
  };

  useEffect(() => {
    if (!props.isOpen) return;

    handleLayerDidMount().catch((error) => {
      console.error("Error in handleLayerDidMount:", error);
    });

    return () => {
      if (cleanupRef.current) {
        cleanupRef.current();
        cleanupRef.current = null;
      }
    };
  }, [props.isOpen]);

  useEffect(() => {
    return () => {
      if (cleanupRef.current) {
        cleanupRef.current();
      }
      
      msalWrapperRef.current = null;
    };
  }, []);

  const styleOptions = useMemo(() => ({
    rootHeight: '600px',
    backgroundColor: 'rgba(255, 255, 255, 0.98)',
    bubbleBackground: '#f3f2f1',
    bubbleFromUserBackground: props.primaryColor || '#0078d4',
    bubbleFromUserTextColor: '#ffffff',
    fontFamily: 'Segoe UI, sans-serif',
    fontSize: '14px',
    sendBoxBackground: '#ffffff',
    sendBoxBorderRadius: '8px',
    sendBoxButtonColor: props.primaryColor || '#0078d4',
    sendBoxButtonColorOnHover: props.accentColor || '#106ebe',
    sendBoxTextColor: '#323130',
    
    markdownRenderHTML: true,
    
    disableFileUpload: !(props.enableFileUpload ?? true), // Enable by default
    uploadAccept: props.supportedFileTypes?.join(',') || 'image/*,.pdf,.doc,.docx,.txt',
    uploadMultiple: true,
    sendAttachmentOn: 'send' as const,
    
    avatarSize: 40,
    botAvatarImage: props.botAvatarImage || '',
    botAvatarInitials: props.botAvatarInitials || props.botName?.charAt(0) || 'AI',
    userAvatarImage: props.userAvatarImage || '',
    userAvatarInitials: props.userAvatarInitials || props.userFriendlyName?.charAt(0) || 'U',
    showAvatarInGroup: 'status' as const,
    
    timestampFormat: 'relative' as const,
    groupTimestamp: 30000,
    timestampColor: '#605e5c',
    showTimestamp: props.showTimestamp ?? true,
    
    typingAnimationDuration: 5000,
    typingAnimationHeight: 20,
    typingAnimationWidth: 64,
    
    suggestedActionLayout: 'stacked' as const,
    suggestedActionsStackedLayoutButtonMaxWidth: 300,
    
    showSpokenText: props.enableSpeech ?? false,
    speechRecognitionContinuous: false,
    
    scrollToEndButtonBehavior: 'unread' as const,
    
    internalLiveRegionFadeAfter: 1000
  }), [props]);

  return (
    <Dialog
      hidden={!props.isOpen}
      onDismiss={props.onDismiss}
      dialogContentProps={dialogContentProps}
      modalProps={modalProps}
    >
      <div className={styles.chatbotContainer}>
        <div className={styles.chatLayout}>
          {showSessionHistory && (
            <div className={styles.sessionHistoryPanel}>
              <div className={styles.sessionHistoryHeader}>
                <span>Chat History</span>
                <IconButton 
                  iconProps={{ iconName: 'ChromeClose' }}
                  onClick={() => setShowSessionHistory(false)}
                  styles={{ 
                    root: { minWidth: '24px', padding: '2px' },
                    icon: { fontSize: '12px' }
                  }}
                />
              </div>
              <div className={styles.sessionHistoryList}>
                {sessionHistory.map(session => (
                  <div 
                    key={session.id}
                    className={`${styles.sessionHistoryItem} ${session.id === sessionId ? styles.activeSession : ''}`}
                    onClick={() => {
                      setSessionId(session.id);
                      setCurrentSessionData(session);
                      localStorage.setItem(currentSessionKey, JSON.stringify(session));
                      // Force WebChat to reinitialize by ending the current DirectLine connection
                      if (directLine) {
                        try {
                          directLine.end();
                        } catch (error) {
                          console.warn('Error ending DirectLine connection for session switch:', error);
                        }
                      }
                    }}
                  >
                    <div className={styles.sessionTitle}>{session.title}</div>
                    <div className={styles.sessionMeta}>
                      {session.messageCount} messages • {new Date(session.timestamp).toLocaleDateString()}
                    </div>
                  </div>
                ))}
                {sessionHistory.length === 0 && (
                  <div className={styles.noSessions}>No previous sessions</div>
                )}
              </div>
            </div>
          )}
          <div className={`${styles.chatMainArea} ${showSessionHistory ? styles.withSidebar : ''}`}>
            {directLine ? (
              <div className={styles.webChatContainer} role="main">
            <FluentThemeProvider>
              <ReactWebChat 
                key={sessionId} // Force reinitialization on session change
                directLine={directLine}
                webSpeechPonyfillFactory={webSpeechPonyfillFactory}
                attachmentMiddleware={attachmentMiddleware}
                store={createStore(
                  currentSessionData ? { activities: currentSessionData.activities } : {},
                  ({ dispatch }: { dispatch: Dispatch }) => (next: any) => (action: any) => {
                  if (props.greet && action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
                    dispatch({
                      meta: { method: "keyboard" },
                      payload: {
                        activity: {
                          channelData: { postBack: true },
                          name: 'startConversation',
                          type: "event"
                        },
                      },
                      type: "DIRECT_LINE/POST_ACTIVITY",
                    });
                    return next(action);
                  }

                  if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
                    const activity = action.payload.activity;
                    if (activity.from?.role === 'bot' && getOAuthCardResourceUri(activity)) {
                      if (directLine) {
                        (directLine as any).postActivity({
                          type: 'invoke',
                          name: 'signin/tokenExchange',
                          value: {
                            id: activity.attachments[0].content.tokenExchangeResource.id,
                            connectionName: activity.attachments[0].content.connectionName,
                            token: directLineToken
                          },
                          from: {
                            id: props.userEmail,
                            name: props.userFriendlyName ?? '',
                            role: "user"
                          }
                        }).subscribe(
                          (id: any) => {
                            if (id === "retry") {
                              return next(action);
                            }
                          },
                          () => {
                            return next(action);
                          }
                        );
                      }
                      return;
                    }
                  }

                  const result = next(action);
                  
                  if (action.type === 'DIRECT_LINE/INCOMING_ACTIVITY' || 
                      action.type === 'DIRECT_LINE/POST_ACTIVITY' ||
                      action.type === 'WEB_CHAT/SEND_MESSAGE') {
                    setTimeout(() => {
                      const state = (next as any).__store?.getState?.() || {};
                      const activities = state.activities || [];
                      if (activities.length > 0) {
                        saveCurrentSession(activities);
                      }
                    }, 500);
                  }
                  
                  return result;
                })}
                styleOptions={styleOptions}
                userID={props.userEmail}
              />
            </FluentThemeProvider>
          </div>
            ) : (
              <div ref={loadingSpinnerRef} className={styles.spinnerContainer}>
                <Spinner label="Loading chatbot..." />
              </div>
            )}
          </div>
        </div>
      </div>
    </Dialog>
  );
};
