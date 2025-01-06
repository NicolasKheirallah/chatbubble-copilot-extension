export interface IChatbotProps {
     botURL: string;
     botName?: string;
     buttonLabel?: string;
     userEmail: string;
     botAvatarImage?: string;
     botAvatarInitials?: string;
     greet?: boolean;
     customScope: string;
     clientID: string;
     authority: string;
     userFriendlyName?: string;
   
     /** Controls whether the chatbot dialog is open */
     isOpen: boolean;
   
     /** Callback invoked when the dialog is dismissed */
     onDismiss: () => void;
   }
   