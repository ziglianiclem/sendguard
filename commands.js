// =====================================================================
//  SendGuard (nouvel Outlook) - Smart Alerts, evenement OnMessageSend (SendMode PromptUser)
// =====================================================================
//  Outlook n'enchaine pas deux invites pendant l'envoi : une seule alerte reprend les destinataires
//   externes (s'il y en a) et l'objet. Boutons :
//    - "Envoyer quand meme" (libelle d'Outlook) = envoi tel quel ;
//    - "Modifier l'objet" = annule l'envoi et ouvre le volet (commandId = bouton sgTaskpane du manifeste),
//      ou l'objet se corrige ; "Valider" y memorise l'instantane et le prochain Envoyer passe sans alerte.
//  Erreur du complement = envoi laisse passer (comme la version COM).

const SG_MAX_SHOWN = 6;     // destinataires listes dans l'alerte (limite d'Outlook : 500 caracteres)

function sgEscapeMd(s) { return s.replace(/[\\`*_{}\[\]()#+\-.!|<>~]/g, "\\$&"); }

function sgMessages(s) {
  const plain = [], md = [];
  const n = s.external.length;
  if (n > 0) {
    const head = n === 1 ? "1 destinataire externe" : n + " destinataires externes";
    plain.push("⚠ " + head + " :");
    md.push("**⚠ " + head + "**");
    for (const e of s.external.slice(0, SG_MAX_SHOWN)) {
      plain.push("  " + e.tag + "  " + sgRecipientLine(e));
      md.push("- " + e.tag + " : " + sgEscapeMd(sgRecipientLine(e)));
    }
    if (n > SG_MAX_SHOWN) { plain.push("  … et " + (n - SG_MAX_SHOWN) + " autres"); md.push("- … et " + (n - SG_MAX_SHOWN) + " autres"); }
  }
  const subj = s.subject.trim();
  if (subj) {
    plain.push("Objet : « " + subj + " »");
    md.push("**Objet :** « " + sgEscapeMd(subj) + " »");
  } else {
    plain.push("⚠ L'objet est vide.");
    md.push("**⚠ L'objet est vide.**");
  }
  const cut = t => t.length <= 500 ? t : t.slice(0, 497) + "...";
  return { plain: cut(plain.join("\n")), md: cut(md.join("\n\n")) };
}

async function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;
    const s = await sgSnapshot(item);
    let confirmed = null;
    try { confirmed = await sgGet(cb => item.sessionData.getAsync(SG_CONFIRMED, cb)); } catch (e) { }
    if (confirmed === s.key) { event.completed({ allowEvent: true }); return; }

    const m = sgMessages(s);
    const options = { allowEvent: false, errorMessage: m.plain, cancelLabel: "Modifier l'objet", commandId: "sgTaskpane" };
    if (Office.context.requirements.isSetSupported("Mailbox", "1.15")) options.errorMessageMarkdown = m.md;
    event.completed(options);
  } catch (e) {
    console.error("SendGuard", e);
    event.completed({ allowEvent: true });
  }
}

Office.onReady();
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
