import { AsyncLocalStorage } from "node:async_hooks";
const requestContext = new AsyncLocalStorage();
function getRequestTokens() {
  return requestContext.getStore();
}
export {
  getRequestTokens,
  requestContext
};
