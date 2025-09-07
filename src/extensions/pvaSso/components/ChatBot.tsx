import * as React from 'react';
import { IChatbotProps } from '../types/IChatBotProps';
import { PVAChatbotDialog } from './PVAChatbotDialog';
import ChatbotErrorBoundary from './ChatbotErrorBoundary';
import styles from '../styles/PvaSsoApplicationCustomizer.module.scss';

const Chatbot: React.FC<IChatbotProps> = (props) => {
  return (
    <ChatbotErrorBoundary fallbackMessage="The chatbot is temporarily unavailable. Please try again.">
      <div className={styles.chatbotWrapper}>
        <PVAChatbotDialog {...props} />
      </div>
    </ChatbotErrorBoundary>
  );
};

export default Chatbot;
