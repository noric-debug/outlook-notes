/* Notes sur le mail — complément Outlook personnel.
 * La note est stockée dans le mail lui-même (propriétés personnalisées du complément),
 * sur la boîte Exchange : rien n'est envoyé ailleurs. */
(function () {
  "use strict";

  var NOTE_KEY = "note";
  var DATE_KEY = "updated";
  var CATEGORY = "📝 Note";
  var SAVE_DELAY_MS = 800;

  var els = {};
  var props = null;        // CustomProperties du mail courant
  var loadToken = 0;       // ignore les réponses d'un mail précédent
  var saveTimer = null;
  var dirty = false;

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) return;
    els.note = document.getElementById("note");
    els.status = document.getElementById("status");
    els.count = document.getElementById("count");
    els.save = document.getElementById("save");
    els.clear = document.getElementById("clear");
    els.subject = document.getElementById("subject");

    els.note.addEventListener("input", onInput);
    els.note.addEventListener("blur", function () { if (dirty) save(); });
    els.save.addEventListener("click", save);
    els.clear.addEventListener("click", clearNote);

    // Panneau épinglé : recharger la note quand on change de mail.
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadCurrentItem);
    }
    loadCurrentItem();
  });

  function loadCurrentItem() {
    clearTimeout(saveTimer);
    dirty = false;
    props = null;
    var token = ++loadToken;
    var item = Office.context.mailbox.item;

    setEnabled(false);
    els.note.value = "";
    updateCount();

    if (!item) {
      els.subject.textContent = "";
      setStatus("Sélectionnez un mail.");
      return;
    }
    els.subject.textContent = item.subject || "(sans objet)";
    setStatus("Chargement…");

    item.loadCustomPropertiesAsync(function (res) {
      if (token !== loadToken) return;
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus("Impossible de lire la note : " + res.error.message, true);
        return;
      }
      props = res.value;
      els.note.value = props.get(NOTE_KEY) || "";
      updateCount();
      setEnabled(true);
      showSavedDate();
    });
  }

  function onInput() {
    dirty = true;
    updateCount();
    setStatus("Modification…");
    clearTimeout(saveTimer);
    saveTimer = setTimeout(save, SAVE_DELAY_MS);
  }

  function save() {
    clearTimeout(saveTimer);
    if (!props) return;
    var token = loadToken;
    var text = els.note.value.replace(/\s+$/, "");
    var hadNote = !!props.get(NOTE_KEY);

    if (text) {
      props.set(NOTE_KEY, text);
      props.set(DATE_KEY, new Date().toISOString());
    } else {
      props.remove(NOTE_KEY);
      props.remove(DATE_KEY);
    }
    dirty = false;
    setStatus("Enregistrement…");

    props.saveAsync(function (res) {
      if (token !== loadToken) return;
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        dirty = true;
        setStatus("Échec de l'enregistrement : " + res.error.message, true);
        return;
      }
      showSavedDate();
      if (text && !hadNote) setCategory(true);
      if (!text && hadNote) setCategory(false);
    });
  }

  function clearNote() {
    if (!els.note.value || !confirm("Supprimer la note de ce mail ?")) return;
    els.note.value = "";
    updateCount();
    save();
  }

  /* Catégorie « 📝 Note » pour repérer les mails annotés dans la liste.
   * Nécessite Mailbox 1.8 ; si indisponible, on ignore silencieusement. */
  function setCategory(on) {
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) return;
    var item = Office.context.mailbox.item;
    if (!on) {
      item.categories.removeAsync([CATEGORY], function () {});
      return;
    }
    ensureMasterCategory(function () {
      item.categories.addAsync([CATEGORY], function () {});
    });
  }

  function ensureMasterCategory(done) {
    var master = Office.context.mailbox.masterCategories;
    master.getAsync(function (res) {
      var exists = res.status === Office.AsyncResultStatus.Succeeded &&
        (res.value || []).some(function (c) { return c.displayName === CATEGORY; });
      if (exists) return done();
      master.addAsync(
        [{ displayName: CATEGORY, color: Office.MailboxEnums.CategoryColor.Preset3 }],
        function () { done(); }
      );
    });
  }

  function showSavedDate() {
    var iso = props && props.get(DATE_KEY);
    if (!iso) { setStatus("Aucune note."); return; }
    var d = new Date(iso);
    setStatus("Enregistrée le " + d.toLocaleDateString("fr-FR") + " à " +
      d.toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" }));
  }

  function updateCount() {
    els.count.textContent = els.note.value.length + " / 2000";
  }

  function setEnabled(on) {
    els.note.disabled = !on;
    els.save.disabled = !on;
    els.clear.disabled = !on;
  }

  function setStatus(msg, isError) {
    els.status.textContent = msg;
    els.status.className = isError ? "error" : "";
  }
})();
