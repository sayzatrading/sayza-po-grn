import { initializeApp } from 'https://www.gstatic.com/firebasejs/12.3.0/firebase-app.js';
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut } from 'https://www.gstatic.com/firebasejs/12.3.0/firebase-auth.js';

const firebaseConfig = {
  apiKey: "AIzaSyBIMinYUij3x7-hzaI7OMiZhgQRz4wkn88",
  authDomain: "sayza-po-grn.firebaseapp.com",
  projectId: "sayza-po-grn",
  storageBucket: "sayza-po-grn.firebasestorage.app",
  messagingSenderId: "964571996539",
  appId: "1:964571996539:web:61c7a646fdc1e3d0e79dde"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
window.firebaseAuth = auth;
window.firebaseSignIn = (email, password) => signInWithEmailAndPassword(auth, email, password);
window.firebaseSignOut = () => signOut(auth);

let currentUser = null;
window.getFirebaseUser = () => currentUser;
window.authReady = new Promise(resolve => {
  onAuthStateChanged(auth, user => {
    currentUser = user;
    document.body.classList.toggle('auth-locked', !user);
    const login = document.getElementById('loginScreen');
    const appShell = document.getElementById('appShell');
    const userEmail = document.getElementById('userEmail');
    if (user) {
      if (login) login.style.display = 'none';
      if (appShell) appShell.style.display = '';
      if (userEmail) userEmail.textContent = user.email || '';
      resolve(user);
      if (user) window.dispatchEvent(new CustomEvent('firebase-auth-signed-in', {detail:user}));
    } else {
      if (login) login.style.display = 'flex';
      if (appShell) appShell.style.display = 'none';
      resolve(null);
    }
  });
});

// Add the Firebase ID token to every API request made by the existing app.
const originalFetch = window.fetch.bind(window);
window.fetch = async (input, init = {}) => {
  const url = typeof input === 'string' ? input : input.url;
  if (url && url.startsWith('/api/')) {
    await window.authReady;
    if (!currentUser) throw new Error('Please sign in first');
    const token = await currentUser.getIdToken();
    const headers = new Headers(init.headers || (input instanceof Request ? input.headers : undefined));
    headers.set('Authorization', `Bearer ${token}`);
    init = {...init, headers};
  }
  return originalFetch(input, init);
};

window.addEventListener('DOMContentLoaded', () => {
  const form = document.getElementById('loginForm');
  const error = document.getElementById('loginError');
  const button = document.getElementById('loginButton');
  if (!form) return;
  form.addEventListener('submit', async e => {
    e.preventDefault();
    error.textContent = '';
    button.disabled = true;
    button.textContent = 'Signing in...';
    try {
      await signInWithEmailAndPassword(auth, document.getElementById('loginEmail').value.trim(), document.getElementById('loginPassword').value);
    } catch (err) {
      error.textContent = err.code === 'auth/invalid-credential' ? 'Invalid email or password.' : (err.message || 'Sign in failed.');
      button.disabled = false;
      button.textContent = 'Sign In';
    }
  });
  const logout = document.getElementById('logoutButton');
  if (logout) logout.addEventListener('click', () => signOut(auth));
});
