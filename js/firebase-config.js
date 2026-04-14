// Centralized Firebase Configuration
// Firebase Storage is removed as per user intent
const firebaseConfig = {
    projectId: "ktu-analysis-backend",
    appId: "1:138350600894:web:b40f59de1071876e3874ef",
    apiKey: "AIzaSyDR8aR6jEpgGXreqtWeEaXEAWdrfRyVxco",
    authDomain: "ktu-analysis-backend.firebaseapp.com",
    messagingSenderId: "138350600894"
};

// Initialize Firebase if not already initialized
if (!firebase.apps.length) {
    firebase.initializeApp(firebaseConfig);
}
