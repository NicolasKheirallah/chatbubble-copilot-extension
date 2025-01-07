import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
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
      aria-label={props.label} // Accessibility enhancement
    >
      <Icon iconName="Chat" className={props.iconClassName} />
    </button>
  );
};

export default ChatToggleButton;
