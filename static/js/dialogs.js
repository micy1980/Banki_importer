/* dialogs.js — drag-and-drop modális ablakok. */

function makeDialogsDraggable() {
  document.querySelectorAll("dialog").forEach(dialog => {
    const handle = dialog.querySelector(".dialog-head");
    if (!handle || handle.dataset.draggableReady) return;
    handle.dataset.draggableReady = "1";
    handle.style.cursor = "move";
    handle.addEventListener("pointerdown", event => {
      if (event.target.closest("button")) return;
      const rect = dialog.getBoundingClientRect();
      const offsetX = event.clientX - rect.left;
      const offsetY = event.clientY - rect.top;
      dialog.style.margin = "0";
      dialog.style.left = `${rect.left}px`;
      dialog.style.top = `${rect.top}px`;
      dialog.style.position = "fixed";
      handle.setPointerCapture(event.pointerId);
      const move = moveEvent => {
        const maxLeft = Math.max(0, window.innerWidth - rect.width);
        const maxTop = Math.max(0, window.innerHeight - 56);
        const left = Math.min(Math.max(0, moveEvent.clientX - offsetX), maxLeft);
        const top = Math.min(Math.max(0, moveEvent.clientY - offsetY), maxTop);
        dialog.style.left = `${left}px`;
        dialog.style.top = `${top}px`;
      };
      const stop = () => {
        handle.removeEventListener("pointermove", move);
        handle.removeEventListener("pointerup", stop);
        handle.removeEventListener("pointercancel", stop);
      };
      handle.addEventListener("pointermove", move);
      handle.addEventListener("pointerup", stop);
      handle.addEventListener("pointercancel", stop);
    });
  });
}
