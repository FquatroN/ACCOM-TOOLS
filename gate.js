const statusEl = document.getElementById("status");
const titleEl = document.getElementById("gate-title");
const descriptionEl = document.getElementById("gate-description");
const emailFieldEl = document.getElementById("email-field");
const passwordFieldEl = document.getElementById("password-field");
const confirmPasswordFieldEl = document.getElementById("confirm-password-field");
const emailEl = document.getElementById("email");
const passwordEl = document.getElementById("password");
const confirmPasswordEl = document.getElementById("confirm-password");
const signinBtn = document.getElementById("signin");
const signupBtn = document.getElementById("signup");
const forgotPasswordBtn = document.getElementById("forgot-password");
const backToSigninBtn = document.getElementById("back-to-signin");

const config = window.APP_CONFIG || {};
const url = String(config.SUPABASE_URL || "").trim();
const key = String(config.SUPABASE_ANON_KEY || "").trim();

const gateState = {
  mode: detectInitialMode(),
  recoverySessionReady: false,
};

if (!window.supabase || !url || !key) {
  setStatus("Configuration missing in config.js", true);
} else {
  const client = window.supabase.createClient(url, key);
  wire(client);
}

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.style.color = isError ? "#b12030" : "#665a53";
}

function detectInitialMode() {
  const search = new URLSearchParams(window.location.search || "");
  const hash = new URLSearchParams(String(window.location.hash || "").replace(/^#/, ""));
  if (search.get("mode") === "recovery" || search.get("type") === "recovery") return "update-password";
  if (hash.get("type") === "recovery") return "update-password";
  return "signin";
}

function redirectRecoveryUrl() {
  return new URL("/gate.html?mode=recovery", window.location.origin).toString();
}

function applyMode(mode) {
  gateState.mode = mode;
  const isSignIn = mode === "signin";
  const isResetRequest = mode === "request-reset";
  const isUpdatePassword = mode === "update-password";

  titleEl.textContent = isUpdatePassword ? "Choose New Password" : isResetRequest ? "Reset Password" : "Sign In";
  descriptionEl.textContent = isUpdatePassword
    ? "Enter your new password to finish the recovery process."
    : isResetRequest
      ? "Enter your email and we will send you a password reset link."
      : "Authenticate first, then you will enter the application.";

  emailFieldEl.hidden = isUpdatePassword;
  passwordFieldEl.hidden = isResetRequest;
  confirmPasswordFieldEl.hidden = !isUpdatePassword;

  signinBtn.textContent = isUpdatePassword ? "Update password" : isResetRequest ? "Send reset link" : "Sign in";
  signupBtn.hidden = !isSignIn;
  forgotPasswordBtn.hidden = !isSignIn;
  backToSigninBtn.hidden = isSignIn;

  if (isSignIn) {
    confirmPasswordEl.value = "";
  } else if (isResetRequest) {
    passwordEl.value = "";
    confirmPasswordEl.value = "";
  }
}

function setLoading(isLoading) {
  [signinBtn, signupBtn, forgotPasswordBtn, backToSigninBtn].forEach((button) => {
    if (button) button.disabled = isLoading;
  });
}

async function handlePrimaryAction(client) {
  if (gateState.mode === "request-reset") {
    await sendResetLink(client);
    return;
  }
  if (gateState.mode === "update-password") {
    await updateRecoveredPassword(client);
    return;
  }
  await signIn(client);
}

async function signIn(client) {
  const email = emailEl.value.trim();
  const password = passwordEl.value.trim();
  if (!email || !password) {
    setStatus("Enter email and password.", true);
    return;
  }
  setLoading(true);
  setStatus("Signing in...");
  const { error } = await client.auth.signInWithPassword({ email, password });
  setLoading(false);
  if (error) {
    setStatus(error.message, true);
    return;
  }
  window.location.href = "/index.html";
}

async function signUp(client) {
  const email = emailEl.value.trim();
  const password = passwordEl.value.trim();
  if (!email || !password) {
    setStatus("Enter email and password.", true);
    return;
  }
  setLoading(true);
  setStatus("Creating account...");
  const { error } = await client.auth.signUp({ email, password });
  setLoading(false);
  if (error) {
    setStatus(error.message, true);
    return;
  }
  setStatus("Account created. You can now sign in.");
}

async function sendResetLink(client) {
  const email = emailEl.value.trim();
  if (!email) {
    setStatus("Enter the email for the password reset link.", true);
    return;
  }
  setLoading(true);
  setStatus("Sending reset link...");
  const { error } = await client.auth.resetPasswordForEmail(email, { redirectTo: redirectRecoveryUrl() });
  setLoading(false);
  if (error) {
    setStatus(error.message, true);
    return;
  }
  setStatus("Password reset email sent. Open the link in the email to choose a new password.");
}

async function updateRecoveredPassword(client) {
  const password = passwordEl.value.trim();
  const confirmPassword = confirmPasswordEl.value.trim();
  if (password.length < 8) {
    setStatus("Password must have at least 8 characters.", true);
    return;
  }
  if (password !== confirmPassword) {
    setStatus("The passwords do not match.", true);
    return;
  }
  setLoading(true);
  setStatus("Updating password...");
  const { error } = await client.auth.updateUser({ password });
  setLoading(false);
  if (error) {
    setStatus(error.message, true);
    return;
  }
  setStatus("Password updated. Redirecting...");
  window.setTimeout(() => {
    window.location.replace("/index.html");
  }, 500);
}

function enableRecoveryModeFromSession(session) {
  if (!session?.user) {
    gateState.recoverySessionReady = false;
    applyMode("update-password");
    setStatus("Open the password reset link from your email to continue.", true);
    return;
  }
  gateState.recoverySessionReady = true;
  applyMode("update-password");
  setStatus(`Resetting password for ${session.user.email || "your account"}.`);
}

async function wire(client) {
  applyMode(gateState.mode);

  signinBtn.addEventListener("click", () => handlePrimaryAction(client));
  signupBtn.addEventListener("click", () => signUp(client));
  forgotPasswordBtn.addEventListener("click", () => {
    applyMode("request-reset");
    setStatus("");
  });
  backToSigninBtn.addEventListener("click", async () => {
    if (gateState.mode === "update-password" || gateState.recoverySessionReady) {
      await client.auth.signOut();
      gateState.recoverySessionReady = false;
    }
    applyMode("signin");
    window.history.replaceState({}, "", "/gate.html");
    setStatus("");
  });
  [emailEl, passwordEl, confirmPasswordEl].forEach((input) => {
    input?.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        handlePrimaryAction(client);
      }
    });
  });

  client.auth.onAuthStateChange((event, session) => {
    if (event === "PASSWORD_RECOVERY") {
      enableRecoveryModeFromSession(session);
      return;
    }
    if (event === "SIGNED_IN" && gateState.mode !== "update-password") {
      window.location.replace("/index.html");
    }
  });

  if (gateState.mode === "update-password") {
    const { data } = await client.auth.getSession();
    enableRecoveryModeFromSession(data?.session || null);
  }
}
