import { useQuery } from '@tanstack/react-query';

export const useSessionCookie = (): 'loading' | 'set' | 'rate-limited' | 'error' => {
  const { isLoading, isError, isFetched, data } = useQuery({
    queryKey: ['session'],
    queryFn: async () => {
      const sudoPassword = localStorage.getItem('equalto.chat.sudo_password');
      if (sudoPassword) {
        const response = await fetch('/sudo', {
          method: 'POST',
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ password: sudoPassword }),
        });

        if (!response.ok) {
          console.error(
            'Sudo password is set, but /sudo endpoint has failed. Ensure password is correct.',
          );
          throw new Error('Fetch call returned response that is not OK.');
        }

        return 'ok';
      }

      const response = await fetch('/session', {
        headers: {
          Accept: 'application/json',
        },
      });

      if (!response.ok) {
        if (response.status === 429) {
          return 'rate-limited';
        }

        throw new Error('Fetch call returned response that is not OK.');
      }

      return 'ok';
    },
    retry: 4,
    staleTime: Infinity,
  });

  if (isLoading) {
    return 'loading';
  }

  if (isError) {
    return 'error';
  }

  if (isFetched) {
    return data === 'ok' ? 'set' : 'rate-limited';
  }

  throw new Error('Neither loading / error / fetched.');
};
