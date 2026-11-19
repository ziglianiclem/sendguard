// =====================================================================
//  SendGuard (nouvel Outlook) - partage entre l'evenement d'envoi (commands.js) et le volet (taskpane.js)
// =====================================================================
//  INTERNE = destinataire de l'annuaire de l'organisation (recipientType user / distributionList),
//   ou adresse du domaine de l'expediteur ou d'un domaine de SG_EXTRA_DOMAINS (sous-domaines compris).
//   externalUser = toujours externe ; other = juge sur le domaine. Adresse illisible = externe.

// Domaines internes en plus de celui de l'expediteur (filiales...). Ex. : ["filiale.com"]
const SG_EXTRA_DOMAINS = [];

// Cle de sessionData : instantane (destinataires + objet) valide dans le volet ; l'envoi suivant
//  passe sans invite tant que rien n'a change.
const SG_CONFIRMED = "sendguardConfirmed";

function sgGet(start) {
  return new Promise((resolve, reject) =>
    start(r => r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error)));
}

function sgDomain(address) {
  if (!address) return null;
  const at = address.lastIndexOf("@");
  return at < 0 || at === address.length - 1 ? null : address.slice(at + 1).trim().toLowerCase();
}

function sgIsInternal(r, domains) {
  const t = r.recipientType;
  if (t === "user" || t === "distributionList") return true;
  if (t === "externalUser") return false;
  for (let s = sgDomain(r.emailAddress); s; ) {
    if (domains.has(s)) return true;
    const dot = s.indexOf(".");
    s = dot < 0 ? null : s.slice(dot + 1);
  }
  return false;
}

// { subject, external: [{tag, name, address}], internal, key }
async function sgSnapshot(item) {
  const [to, cc, bcc, subject] = await Promise.all([
    sgGet(cb => item.to.getAsync(cb)),
    sgGet(cb => item.cc.getAsync(cb)),
    sgGet(cb => item.bcc.getAsync(cb)),
    sgGet(cb => item.subject.getAsync(cb)),
  ]);
  let from = Office.context.mailbox.userProfile.emailAddress;
  try { const f = await sgGet(cb => item.from.getAsync(cb)); if (f && f.emailAddress) from = f.emailAddress; } catch (e) { }

  const domains = new Set(SG_EXTRA_DOMAINS.map(d => d.toLowerCase().replace(/^@/, "")));
  if (sgDomain(from)) domains.add(sgDomain(from));

  const all = [], external = [];
  let internal = 0;
  for (const [tag, list] of [["À", to], ["Cc", cc], ["Cci", bcc]])
    for (const r of list || []) {
      all.push(tag + ":" + (r.emailAddress || r.displayName || "").toLowerCase());
      if (sgIsInternal(r, domains)) internal++;
      else external.push({ tag, name: r.displayName || "", address: r.emailAddress || "" });
    }
  all.sort();
  return { subject: subject || "", external, internal, key: JSON.stringify([all, subject || ""]) };
}

function sgRecipientLine(e) {
  const addr = e.address && e.address.includes("@") ? e.address : "(adresse non reconnue)";
  return !e.name || e.name.toLowerCase() === addr.toLowerCase() ? addr : e.name + " <" + addr + ">";
}
