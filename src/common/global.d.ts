// TODO: Refactor or eliminate
import { IEpmodern } from "./ep";
declare global {
  interface Window {
    Epmodern: IEpmodern;
  }
}
//window.EP = {};
