import { useCallback, useState } from 'react';

type MailSubmitState =
  | { state: 'waiting' }
  | { state: 'submitting' }
  | { state: 'success' }
  | { state: 'error' };

export function useEmailSubmit(): {
  submitState: MailSubmitState;
  setMail: (mail: string) => void;
  submitMail: () => void;
} {
  const [submitState, setSubmitState] = useState<MailSubmitState>({ state: 'waiting' });
  const [mail, setMail] = useState('');
  const submitMail = useCallback(() => {
    setSubmitState({ state: 'submitting' });

    const simpleEmailRegex = /^[^@\s]+@[^@\s]+$/;
    const sanitizedMail = mail.trim();

    if (simpleEmailRegex.test(sanitizedMail)) {
      submitMail(sanitizedMail);
    } else {
      setSubmitState({ state: 'error' });
    }

    async function submitMail(sanitizedMail: string) {
      try {
        const response = await fetch('./signup', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Accept: 'application/json',
          },
          body: JSON.stringify({ email: sanitizedMail }),
        });

        if (!response.ok) {
          setSubmitState({ state: 'error' });
          return;
        }

        setSubmitState({ state: 'success' });
      } catch {
        setSubmitState({ state: 'error' });
      }
    }
  }, [mail]);

  return { submitState, setMail, submitMail };
}
