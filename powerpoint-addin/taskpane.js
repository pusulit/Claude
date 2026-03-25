/* global Office, PowerPoint */

Office.onReady(function (info) {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("apply-btn").addEventListener("click", applyRoundedCorners);
  }
});

/**
 * Legge il raggio degli angoli arrotondati dalla prima forma selezionata
 * e lo applica a tutte le altre forme selezionate.
 *
 * Richiede PowerPoint JavaScript API requirement set 1.5+ (per getSelectedShapes)
 * e 1.6+ (per shape.adjustments).
 */
async function applyRoundedCorners() {
  const btn = document.getElementById("apply-btn");
  btn.disabled = true;
  showStatus("Elaborazione in corso...", "info");

  try {
    await PowerPoint.run(async (context) => {
      // Recupera le forme selezionate
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      const shapes = selectedShapes.items;

      if (shapes.length < 2) {
        showStatus("Seleziona almeno 2 forme.", "error");
        return;
      }

      // Carica il tipo geometrico di tutte le forme
      shapes.forEach((shape) => shape.load("geometricShapeType, name"));
      await context.sync();

      // Verifica che la prima forma sia un rettangolo arrotondato
      const firstShape = shapes[0];
      if (firstShape.geometricShapeType !== PowerPoint.GeometricShapeType.roundRectangle) {
        showStatus(
          `La prima forma selezionata ("${firstShape.name}") non è un rettangolo arrotondato.`,
          "error"
        );
        return;
      }

      // Carica gli adjustment della prima forma per leggere il raggio
      firstShape.adjustments.load("items");
      await context.sync();

      if (!firstShape.adjustments.items.length) {
        showStatus("Impossibile leggere il raggio della prima forma.", "error");
        return;
      }

      // Il primo adjustment di un roundRectangle è il raggio degli angoli
      // Il valore è un numero tra 0 (nessun arrotondamento) e 0.5 (massimo arrotondamento)
      const cornerRadius = firstShape.adjustments.items[0].value;

      // Applica il raggio a tutte le forme successive che sono rettangoli arrotondati
      let appliedCount = 0;
      let skippedCount = 0;

      for (let i = 1; i < shapes.length; i++) {
        const shape = shapes[i];
        if (shape.geometricShapeType === PowerPoint.GeometricShapeType.roundRectangle) {
          shape.adjustments.load("items");
          await context.sync();

          if (shape.adjustments.items.length > 0) {
            shape.adjustments.items[0].value = cornerRadius;
            appliedCount++;
          } else {
            skippedCount++;
          }
        } else {
          skippedCount++;
        }
      }

      await context.sync();

      // Messaggio di risultato
      if (appliedCount === 0) {
        showStatus(
          "Nessuna forma arrotondata trovata tra le forme selezionate (esclusa la prima).",
          "error"
        );
      } else if (skippedCount > 0) {
        showStatus(
          `Angoli uniformati su ${appliedCount} forma/e. ${skippedCount} forma/e ignorata/e (non rettangoli arrotondati).`,
          "success"
        );
      } else {
        showStatus(
          `Angoli arrotondati uniformati su ${appliedCount} forma/e.`,
          "success"
        );
      }
    });
  } catch (error) {
    if (error instanceof PowerPoint.RichApiException) {
      // API non disponibile (requirement set non soddisfatto)
      if (error.code === PowerPoint.ErrorCodes.apiNotFound) {
        showStatus(
          "Questa funzione richiede una versione più recente di PowerPoint (Microsoft 365).",
          "error"
        );
      } else {
        showStatus(`Errore PowerPoint: ${error.message}`, "error");
      }
    } else {
      showStatus(`Errore: ${error.message || error}`, "error");
    }
    console.error("applyRoundedCorners error:", error);
  } finally {
    btn.disabled = false;
  }
}

/**
 * Mostra un messaggio di stato nel pannello.
 * @param {string} message - Il testo da mostrare
 * @param {"success"|"error"|"info"} type - Il tipo di messaggio
 */
function showStatus(message, type) {
  const statusEl = document.getElementById("status");
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
}
