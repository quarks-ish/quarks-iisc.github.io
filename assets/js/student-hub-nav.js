// Navbar "Student Hub" link.
//   Microsoft session exists -> go to the hub (the hub page then asks Supabase).
//   No session               -> start the Microsoft login; it returns to the hub.
// MSAL is only downloaded when the link is clicked.
//
// Also owns the SweetAlert2 popups. Login and logout are full-page redirects,
// so a popup meant for the next page is stored in sessionStorage ("flash")
// and shown when that page loads.
(function () {
  var FLASH_KEY = "student-hub-flash";
  var PENDING_KEY = "student-hub-login-pending";

  function store(key, value) {
    try { sessionStorage.setItem(key, value); } catch (e) {}
  }
  function read(key) {
    try { return sessionStorage.getItem(key); } catch (e) { return null; }
  }
  function drop(key) {
    try { sessionStorage.removeItem(key); } catch (e) {}
  }

  function show(type, title, text) {
    if (!window.Swal) return;
    var dark = document.body.classList.contains("dark-mode") ||
      document.documentElement.classList.contains("dark-mode");
    window.Swal.fire({
      icon: type, // "success" | "error"
      title: title,
      text: text,
      confirmButtonText: "OK",
      confirmButtonColor: "#343a40",
      timer: type === "success" ? 2500 : undefined,
      timerProgressBar: type === "success",
      background: dark ? "#1e1e1e" : undefined,
      color: dark ? "#e0e0e0" : undefined,
    });
  }

  // Used by the hub page (student-hub/index.html).
  window.StudentHubAlert = {
    show: show,
    // Show on the next page load (use right before a redirect).
    flash: function (type, title, text) {
      store(FLASH_KEY, JSON.stringify({ type: type, title: title, text: text }));
    },
    // Set when a login starts; the hub uses it to report success or failure.
    markLoginPending: function () { store(PENDING_KEY, "1"); },
    takeLoginPending: function () {
      var pending = read(PENDING_KEY) === "1";
      drop(PENDING_KEY);
      return pending;
    },
  };

  // Show a popup left by the previous page, if any.
  var flash = read(FLASH_KEY);
  if (flash) {
    drop(FLASH_KEY);
    try {
      var f = JSON.parse(flash);
      show(f.type, f.title, f.text);
    } catch (e) {}
  }

  var link = document.getElementById("student-hub-link");
  if (!link) return;

  link.addEventListener("click", function (event) {
    // Leave ctrl/cmd/middle-click (open in new tab) alone: the hub page
    // checks the session itself.
    if (event.ctrlKey || event.metaKey || event.shiftKey || event.button !== 0) return;
    event.preventDefault();
    if (link.getAttribute("aria-busy") === "true") return;
    link.setAttribute("aria-busy", "true");
    window.StudentHubAlert.markLoginPending();

    import(link.getAttribute("data-auth-module"))
      .then(function (auth) {
        return auth.getMicrosoftIdToken().then(function (idToken) {
          if (idToken) {
            window.location.href = link.href;
          } else {
            return auth.startMicrosoftLogin();
          }
        });
      })
      .catch(function (err) {
        link.removeAttribute("aria-busy");
        // Sign-in code couldn't load. Go to the hub anyway: it fails closed
        // and sends the visitor back to the home page.
        console.error("Student Hub: sign-in unavailable", err);
        window.location.href = link.href;
      });
  });
})();
