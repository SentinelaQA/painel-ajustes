import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

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
export const db = getFirestore(app);

// App Firebase secundário, usado só pra criação de novos usuários na aba de Admin.
// O SDK do Firebase Auth troca a sessão ativa pro usuário recém-criado assim que ele
// é criado — com uma instância separada isso não derruba a sessão de quem está
// logado como admin no app principal.
const secondaryApp = initializeApp(firebaseConfig, "secondary");
export const secondaryAuth = getAuth(secondaryApp);

export default app;
