// PvaSsoApplicationCustomizer.tsx
import React from 'react';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import * as ReactDOM from 'react-dom';
import * as strings from 'PvaSsoApplicationCustomizerStrings';
import { override } from '@microsoft/decorators';
import { initializeIcons, getTheme } from '@fluentui/react';
import { Icon } from '@fluentui/react/lib/Icon';

import styles from './PvaSsoApplicationCustomizer.module.scss';
import { IChatbotProps } from './components/IChatBotProps';
import Chatbot from './components/ChatBot';

initializeIcons();

const LOG_SOURCE: string = 'PvaSsoApplicationCustomizer';

interface IChatState {
  isOpen: boolean;
}

interface IChatToggleButtonProps {
  label: string;
  onClick: () => void;
  className?: string;
  iconClassName?: string;
}

const ChatToggleButton: React.FC<IChatToggleButtonProps> = (props) => {
  return (
    <button
      className={props.className}
      title={props.label}
      onClick={props.onClick}
      aria-label={props.label}
    >
      <Icon iconName="Chat" className={props.iconClassName} />
    </button>
  );
};

export interface IPvaSsoApplicationCustomizerProperties {
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
}

export default class PvaSsoApplicationCustomizer
  extends BaseApplicationCustomizer<IPvaSsoApplicationCustomizerProperties> {

  private _buttonPlaceholder: PlaceholderContent | undefined;
  private _chatPlaceholder: PlaceholderContent | undefined;
  private _chatState: IChatState = { isOpen: false };

  @override
  public async onInit(): Promise<void> {
    try {
      Log.info(LOG_SOURCE, 'Initializing application customizer');
      
      this._setDefaultProperties();
      this._applyThemeColor();

      // Create both placeholders
      this._buttonPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom
      );

      this._chatPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom
      );

      this._renderPlaceholders();

      return Promise.resolve();
    } catch (error) {
      Log.error(LOG_SOURCE, error);
      return Promise.reject(error);
    }
  }

  private _setDefaultProperties(): void {
    if (!this.properties.buttonLabel || this.properties.buttonLabel.trim() === '') {
      this.properties.buttonLabel = strings.DefaultButtonLabel || 'Chat';
    }
    if (!this.properties.botName || this.properties.botName.trim() === '') {
      this.properties.botName = strings.DefaultBotName || 'Support Chat';
    }
    if (this.properties.greet !== true) {
      this.properties.greet = false;
    }
  }

  private _applyThemeColor(): void {
    const theme = getTheme();
    const styleEl = document.createElement('style');
    styleEl.innerHTML = `
      .${styles.modernChatButton} {
        background-color: ${theme.palette.themePrimary};
        color: ${theme.palette.white};
      }
      .${styles.modernChatButton}:hover {
        background-color: ${theme.palette.themeDarkAlt};
      }
    `;
    document.head.appendChild(styleEl);
  }

  private _renderPlaceholders(): void {
    // Render button
    if (this._buttonPlaceholder && !this._buttonPlaceholder.isDisposed) {
      this._buttonPlaceholder.domElement.className = styles.modernChatContainer;
      ReactDOM.render(
        <ChatToggleButton
          label={this.properties.buttonLabel || 'Chat'}
          onClick={this._toggleChat}
          className={styles.modernChatButton}
          iconClassName={styles.modernChatIcon}
        />,
        this._buttonPlaceholder.domElement
      );
    }

    // Render chat if open
    if (this._chatPlaceholder && !this._chatPlaceholder.isDisposed) {
      if (this._chatState.isOpen) {
        const user = this.context.pageContext.user;
        const userEmail: string = user?.email ?? '';
        const userFriendlyName: string = user?.displayName ?? '';

        const chatbotProps: IChatbotProps = {
          botURL: this.properties.botURL,
          botName: this.properties.botName,
          buttonLabel: this.properties.buttonLabel,
          userEmail,
          botAvatarImage: this.properties.botAvatarImage,
          botAvatarInitials: this.properties.botAvatarInitials,
          greet: this.properties.greet,
          customScope: this.properties.customScope,
          clientID: this.properties.clientID,
          authority: this.properties.authority,
          userFriendlyName,
          isOpen: true,
          onDismiss: this._hideChat
        };

        ReactDOM.render(<Chatbot {...chatbotProps} />, this._chatPlaceholder.domElement);
      } else {
        ReactDOM.unmountComponentAtNode(this._chatPlaceholder.domElement);
      }
    }
  }

  private _toggleChat = (): void => {
    if (!this.properties.botURL?.trim()) {
      Log.error(LOG_SOURCE, new Error('Error: botURL is undefined or empty'));
      return;
    }

    this._chatState.isOpen = !this._chatState.isOpen;
    this._renderPlaceholders();
  };

  private _hideChat = (): void => {
    this._chatState.isOpen = false;
    this._renderPlaceholders();
  };

  protected onDispose(): void {
    if (this._buttonPlaceholder && !this._buttonPlaceholder.isDisposed) {
      ReactDOM.unmountComponentAtNode(this._buttonPlaceholder.domElement);
    }
    if (this._chatPlaceholder && !this._chatPlaceholder.isDisposed) {
      ReactDOM.unmountComponentAtNode(this._chatPlaceholder.domElement);
    }
  }
}