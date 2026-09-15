Office.onReady(() => {
  const input = document.getElementById("logInput");
  input.focus();

  input.addEventListener("keypress", (e) => {
    if (e.key === "Enter") submit();
  });

  // Status messages sent back from commands.js after each write.
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
    const status = arg.message;
    const isError = status.toLowerCase().indexOf("error") === 0 || status.toLowerCase().indexOf("error:") !== -1;
    setStatus(status, isError ? "err" : "ok");

    input.disabled = false;
    input.value = "";
    input.focus();
  });
});

function submit() {
  const input = document.getElementById("logInput");
  const text = input.value.trim();
  if (text === "") return;

  input.disabled = true;
  setStatus("Logging…", "busy");
  Office.context.ui.messageParent(text);
}

function setStatus(message, kind) {
  const el = document.getElementById("status");
  el.textContent = message;
  el.className = kind === "ok" ? "status-ok" : kind === "err" ? "status-err" : "status-busy";
}
