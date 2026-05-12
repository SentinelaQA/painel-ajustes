import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";

const firebaseConfig = {
  apiKey: "AIzaSyC2VPwlPbrOfxsGnkRpV-oziMaOuAlP_Xs",
  authDomain: "ajustes-85c9c.firebaseapp.com",
  projectId: "ajustes-85c9c",
  storageBucket: "ajustes-85c9c.firebasestorage.app",
  messagingSenderId: "358429411322",
  appId: "1:358429411322:web:bb9046ccb914fe847e8643",
  measurementId: "G-RRSRC8H89N",
};

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export default app;
