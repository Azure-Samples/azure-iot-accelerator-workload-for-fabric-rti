import { openBrowser, signIn } from "./auth.js";
import { authConfig } from "./config.js";

const config = authConfig();
const context = await openBrowser(config);
try {
  await signIn(context.pages()[0] ?? await context.newPage(), config);
} finally {
  await context.close();
}
