import { EP, IEpModern } from "./ep";

declare global {
  interface Window {
    ElevatePoint:EP;
    EpModern: IEpModern;
  }
}

