// =====================================================================
//  SendGuard (nouvel Outlook) - volet : destinataires externes + objet modifiable
// =====================================================================
//  "Valider" ecrit l'objet, puis memorise l'instantane (destinataires + objet) dans sessionData :
//   le prochain clic sur Envoyer part sans alerte. Toute modification ulterieure redeclenche l'alerte.

Office.onReady(async () => {
  const item = Office.context.mailbox.item;
  const box = document.getElementById("subject");
  const btn = document.getElementById("validate");
  const status = document.getElementById("status");
  const say = (text, cls) => { status.textContent = text; status.className = cls || "muted"; };

  async function refresh() {
    const s = await sgSnapshot(item);
    const sec = document.getElementById("recipients");
    sec.replaceChildren();
    const h = document.createElement("h2");
    if (s.external.length === 0) {
      h.textContent = "Aucun destinataire externe";
      sec.append(h);
    } else {
      h.className = "warn";
      h.textContent = "⚠ " + (s.external.length === 1 ? "1 destinataire externe" : s.external.length + " destinataires externes");
      const ul = document.createElement("ul");
      for (const e of s.external) {
        const li = document.createElement("li");
        li.textContent = e.tag + " : " + sgRecipientLine(e);
        ul.append(li);
      }
      sec.append(h, ul);
    }
    if (s.internal > 0) {
      const p = document.createElement("div");
      p.className = "muted";
      p.textContent = s.internal === 1 ? "+ 1 interne" : "+ " + s.internal + " internes";
      sec.append(p);
    }
    return s;
  }

  try {
    const s = await refresh();
    box.value = s.subject;
    btn.disabled = !box.value.trim();
    box.focus();
    box.setSelectionRange(box.value.length, box.value.length);
  } catch (e) { say("Lecture impossible : " + (e && e.message || e), "err"); }

  box.addEventListener("input", () => { btn.disabled = !box.value.trim(); say(""); });
  box.addEventListener("keydown", e => { if (e.key === "Enter" && !btn.disabled) btn.click(); });

  btn.addEventListener("click", async () => {
    btn.disabled = true;
    try {
      const subject = box.value.trim();
      await sgGet(cb => item.subject.setAsync(subject, cb));
      const s = await refresh();     // destinataires eventuellement modifies entre-temps
      await sgGet(cb => item.sessionData.setAsync(SG_CONFIRMED, s.key, cb));
      say("Validé. Cliquez sur Envoyer : le mail partira sans nouvelle alerte.", "ok");
    } catch (e) {
      say("Échec : " + (e && e.message || e), "err");
    }
    btn.disabled = !box.value.trim();
  });
});
