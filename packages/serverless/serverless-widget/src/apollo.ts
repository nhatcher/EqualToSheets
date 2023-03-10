import { ApolloClient, InMemoryCache, createHttpLink } from '@apollo/client';
import { setContext } from '@apollo/client/link/context';
import { getGraphqlEndpoint } from './domain';
import { getLicenseKey } from './license';

/**
 * NB: We want to create client at the moment of rendering provider -> we want to allow to change `getGraphqlEndpoint` at runtime
 */
export const getApollo = () => {
  const httpLink = createHttpLink({
    uri: getGraphqlEndpoint(),
  });

  const authLink = setContext((_, { headers }) => {
    return {
      headers: {
        ...headers,
        authorization: getLicenseKey() ? `Bearer ${getLicenseKey()}` : "",
      }
    }
  });

  return new ApolloClient({
    link: authLink.concat(httpLink),
    cache: new InMemoryCache(),
  });
};
