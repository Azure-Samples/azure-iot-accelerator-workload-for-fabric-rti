import { keepSystemAwake } from "../src/awake.js";

export default async function setup() {
  return keepSystemAwake();
}
