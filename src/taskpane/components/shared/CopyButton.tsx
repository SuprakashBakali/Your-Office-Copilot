import React, { useState } from 'react';
import { Button, Tooltip } from '@fluentui/react-components';
import { Copy16Regular, Checkmark16Regular } from '@fluentui/react-icons';

interface CopyButtonProps {
  text?: string;
  textToCopy?: string;
}

export const CopyButton: React.FC<CopyButtonProps> = React.memo(({ text, textToCopy }) => {
  const [copied, setCopied] = useState(false);
  const content = text || textToCopy || '';

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(content);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      // Fallback for older browsers / restricted contexts
      const textarea = document.createElement('textarea');
      textarea.value = content;
      textarea.style.position = 'fixed';
      textarea.style.opacity = '0';
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  return (
    <Tooltip content={copied ? 'Copied!' : 'Copy to clipboard'} relationship="label">
      <Button
        appearance="transparent"
        icon={copied ? <Checkmark16Regular style={{ color: '#76B900' }} /> : <Copy16Regular />}
        size="small"
        onClick={handleCopy}
      />
    </Tooltip>
  );
});
