// Used only if this repository is deployed as a Cloudflare Pages project (instead of a Worker).
// Pages then needs a D1 binding called DB under Settings → Bindings.
import { handleApi } from '../../worker/api.js';

export const onRequest = (context) => handleApi(context.request, context.env);
