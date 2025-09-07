export interface IChatbotProps {
  /** The URL endpoint for the bot */
  botURL: string;
  
  /** Display name for the bot */
  botName?: string;
  
  /** Email of the current user */
  userEmail: string;
  
  /** Display name of the current user */
  userFriendlyName?: string;
  
  /** Custom OAuth scope */
  customScope: string;
  
  /** Azure AD client ID */
  clientID: string;
  
  /** Azure AD authority URL */
  authority: string;
  
  /** SharePoint context */
  context: any;

  /** Controls whether the chatbot dialog is open */
  isOpen: boolean;
 
  /** Callback invoked when the dialog is dismissed */
  onDismiss: () => void;

  // Avatar Configuration
  /** URL for bot's avatar image */
  botAvatarImage?: string;
  
  /** Initials to display in bot's avatar */
  botAvatarInitials?: string;
  
  /** URL for user's avatar image */
  userAvatarImage?: string;
  
  /** Initials to display in user's avatar */
  userAvatarInitials?: string;

  // Feature Toggles
  /** Whether to send initial greeting message */
  greet?: boolean;
  
  /** Enable file upload functionality */
  enableFileUpload?: boolean;
  
  /** Enable speech-to-text functionality */
  enableSpeech?: boolean;
  
  /** Enable text-to-speech functionality */
  enableSpeechSynthesis?: boolean;
  
  /** Show message timestamps */
  showTimestamp?: boolean;
  
  /** Enable typing indicators */
  sendTypingIndicator?: boolean;
  
  /** Enable adaptive cards support */
  enableAdaptiveCards?: boolean;

  // Customization Options
  /** Custom theme colors */
  primaryColor?: string;
  
  /** Custom accent color */
  accentColor?: string;
  
  /** Maximum file upload size in MB */
  maxUploadSizeMB?: number;
  
  /** Supported file types for upload */
  supportedFileTypes?: string[];
  
  /** Enable session persistence */
  enableSessionPersistence?: boolean;
  
  /** Session storage duration in hours (default: 24) */
  sessionDurationHours?: number;
}