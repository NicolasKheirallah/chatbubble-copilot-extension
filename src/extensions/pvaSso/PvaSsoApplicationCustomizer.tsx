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
import { PVAChatbotDialog } from './components/PVAChatbotDialog';
import { ConfigurationService, IChatbotConfiguration } from './services/ConfigurationService';

initializeIcons();

const LOG_SOURCE: string = 'PvaSsoApplicationCustomizer';

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
    >
      <Icon iconName="Chat" className={props.iconClassName} />
    </button>
  );
};

export interface IPvaSsoApplicationCustomizerProperties extends IChatbotConfiguration {}

export default class PvaSsoApplicationCustomizer
  extends BaseApplicationCustomizer<IPvaSsoApplicationCustomizerProperties> {

  private _placeholder: PlaceholderContent | undefined;
  private _chatVisible: boolean = false;
  private _configurationService: ConfigurationService;
  private _configuration: IChatbotConfiguration | undefined;

  @override
  public async onInit(): Promise<void> {
    try {
      Log.info(LOG_SOURCE, 'Initializing application customizer');

      // Initialize configuration service
      this._configurationService = new ConfigurationService(this.context);
      
      // Get configuration
      this._configuration = await this._configurationService.getConfiguration();
      
      Log.info(LOG_SOURCE, `Configuration loaded: ${JSON.stringify(this._configuration)}`);

      this._setDefaultProperties();
      this._applyThemeColor();

      // Create placeholder
      this._placeholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDispose }
      );

      if (!this._placeholder) {
        console.error('Could not find placeholder');
        return Promise.reject(new Error('Could not find placeholder'));
      }

      this._renderPlaceholders();

      return Promise.resolve();
    } catch (error) {
      Log.error(LOG_SOURCE, error);
      console.error('Failed to initialize application customizer:', error);
      return Promise.reject(error);
    }
  }

  private _setDefaultProperties(): void {
    if (!this._configuration) return;

    if (!this._configuration.buttonLabel || this._configuration.buttonLabel.trim() === '') {
      this._configuration.buttonLabel = strings.DefaultButtonLabel || 'Chat';
    }
    if (!this._configuration.botName || this._configuration.botName.trim() === '') {
      this._configuration.botName = strings.DefaultBotName || 'Support Chat';
    }
    if (this._configuration.greet !== true) {
      this._configuration.greet = false;
    }
  }

  private _applyThemeColor(): void {
    const theme = getTheme();
    const styleEl = document.createElement('style');
    styleEl.innerHTML = `
      .${styles.modernChatButton} {
        background-color: ${theme.palette.themePrimary};
        color: ${theme.palette.white};
        border: none;
        padding: 0;
        border-radius: 50%;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 2px 6px rgba(0,0,0,0.3);
        width: 60px;
        height: 60px;
        transition: background-color 0.3s ease;
      }
      .${styles.modernChatButton}:hover {
        background-color: ${theme.palette.themeDarkAlt};
      }
    `;
    document.head.appendChild(styleEl);
  }

  private _renderPlaceholders(): void {
    if (this._placeholder && !this._placeholder.isDisposed) {
      this._placeholder.domElement.className = styles.modernChatContainer;

      if (this._chatVisible && this._configuration) {
        // Render chat dialog
        const user = this.context.pageContext.user;
        const chatbotProps: IChatbotProps = {
          ...this._configuration,
          context: this.context,
          userEmail: user?.email ?? '',
          userFriendlyName: user?.displayName ?? '',
          isOpen: true,
          onDismiss: this._hideChat
        };

        ReactDOM.render(<PVAChatbotDialog {...chatbotProps} />, this._placeholder.domElement);
      } else {
        // Render toggle button
        ReactDOM.render(
          <ChatToggleButton
            label={this._configuration?.buttonLabel || 'Chat'}
            onClick={this._toggleChat}
            className={styles.modernChatButton}
            iconClassName={styles.modernChatIcon}
          />,
          this._placeholder.domElement
        );
      }
    }
  }

  private _toggleChat = (): void => {
    if (!this._configuration) {
      console.error('Configuration not available');
      return;
    }

    this._chatVisible = !this._chatVisible;
    this._renderPlaceholders();
  };

  private _hideChat = (): void => {
    this._chatVisible = false;
    this._renderPlaceholders();
  };

  private _onDispose = (): void => {
    if (this._placeholder && !this._placeholder.isDisposed) {
      ReactDOM.unmountComponentAtNode(this._placeholder.domElement);
    }
  };
}
