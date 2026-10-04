#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sociogramma.py — Genera un diagramma di Moreno (sociogramma) a partire da un file
Excel con le colonne: ID, COGNOME, NOME, SI-1, SI-2, SI-3, NO-1, NO-2.

Ogni riga è una persona. SI-1, SI-2, SI-3 contengono, in ordine di preferenza,
gli ID delle persone con cui vorrebbe lavorare; NO-1 e NO-2 gli ID delle persone
con cui non vorrebbe lavorare. Nei grafici le persone sono indicate con nome e
cognome (COGNOME e NOME sono facoltative: se mancano si usa l'ID).

Output prodotti (cartella indicata con -o):
  sociogramma_scelte.png     grafico delle sole scelte
  sociogramma_rifiuti.png    grafico dei soli rifiuti
  sociogramma_completo.png   opzionale (--insieme), scelte e rifiuti sovrapposti
  sociogramma.html           versione interattiva offline (zoom, nodi trascinabili)
  indici.xlsx                indici individuali e di gruppo, matrice, anomalie
  report.docx                report commentato (richiede python-docx)
  report.pdf                 opzionale (--pdf), conversione via LibreOffice

Esempi:
  python sociogramma.py dati.xlsx
  python sociogramma.py dati.xlsx -o risultati --formato pdf --insieme --pdf
  python sociogramma.py --crea-esempio esempio.xlsx

Dipendenze: pandas, numpy, networkx, matplotlib, openpyxl, python-docx (facoltativa)
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import shutil
import subprocess
import sys
from pathlib import Path

PACCHETTI = ["pandas", "numpy", "networkx", "matplotlib", "openpyxl", "python-docx"]


def _importa_librerie() -> None:
    """Importa le librerie esterne rendendole disponibili a tutto il modulo."""
    global np, pd, nx, matplotlib, plt, Line2D, Circle
    import numpy as _np
    import pandas as _pd
    import networkx as _nx
    import matplotlib as _mpl
    _mpl.use("Agg")
    import matplotlib.pyplot as _plt
    from matplotlib.lines import Line2D as _Line2D
    from matplotlib.patches import Circle as _Circle
    np, pd, nx, matplotlib, plt = _np, _pd, _nx, _mpl, _plt
    Line2D, Circle = _Line2D, _Circle


def _installa_pacchetti() -> tuple[bool, str]:
    """Installa le librerie mancanti con pip; ritorna (riuscito, messaggi)."""
    opzioni = {}
    if sys.platform.startswith("win"):
        opzioni["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    for extra in (["--user"], []):
        esito = subprocess.run([sys.executable, "-m", "pip", "install", *extra, *PACCHETTI],
                               capture_output=True, text=True, **opzioni)
        if esito.returncode == 0:
            return True, esito.stdout[-1500:]
    return False, ((esito.stdout or "") + (esito.stderr or ""))[-1500:]


def _finestra_installazione() -> bool:
    """Chiede conferma e installa mostrando una finestra di attesa."""
    try:
        import queue
        import threading
        import tkinter as tk
        from tkinter import messagebox, ttk
    except Exception:
        return False

    try:
        radice = tk.Tk()
        radice.withdraw()
    except Exception:
        return False        # nessuna interfaccia grafica disponibile
    domanda = ("Per funzionare il programma ha bisogno di alcuni componenti che "
               "non sono ancora installati su questo computer.\n\n"
               "Vuoi installarli adesso?\n\n"
               "Serve la connessione a internet e possono volerci alcuni minuti. "
               "Va fatto una volta sola.")
    if not messagebox.askyesno("Prima configurazione", domanda):
        radice.destroy()
        return False

    attesa = tk.Toplevel(radice)
    attesa.title("Installazione in corso")
    attesa.geometry("430x130")
    attesa.resizable(False, False)
    attesa.protocol("WM_DELETE_WINDOW", lambda: None)
    ttk.Label(attesa, padding=16, justify="left",
              text="Installazione dei componenti in corso.\n"
                   "Non chiudere questa finestra: al termine il programma si apre "
                   "da solo.").pack()
    barra = ttk.Progressbar(attesa, mode="indeterminate", length=380)
    barra.pack(pady=6)
    barra.start(12)
    attesa.deiconify()
    attesa.update()

    risultati: "queue.Queue[tuple[bool, str]]" = queue.Queue()
    threading.Thread(target=lambda: risultati.put(_installa_pacchetti()), daemon=True).start()

    esito = {}

    def controlla():
        try:
            esito["dati"] = risultati.get_nowait()
            attesa.destroy()
            radice.quit()
        except queue.Empty:
            radice.after(200, controlla)

    radice.after(200, controlla)
    radice.mainloop()
    riuscito, messaggi = esito.get("dati", (False, ""))
    if not riuscito:
        messagebox.showerror(
            "Installazione non riuscita",
            "Non sono riuscito a installare i componenti.\n\n"
            "Verifica la connessione a internet e riprova. Se il computer è "
            "aziendale, potrebbe servire l'aiuto dell'assistenza informatica.\n\n"
            f"Dettaglio tecnico:\n{messaggi[-600:]}")
    radice.destroy()
    return riuscito


def _componenti_mancanti(errore) -> None:
    """Prova a installare da sola le librerie mancanti, altrimenti spiega come fare."""
    if _finestra_installazione():
        try:
            _importa_librerie()
            return
        except ImportError:
            try:
                import tkinter
                from tkinter import messagebox
                radice = tkinter.Tk()
                radice.withdraw()
                messagebox.showinfo("Quasi fatto",
                                    "Installazione completata. Chiudi e riapri il "
                                    "programma per usarlo.")
                radice.destroy()
            except Exception:
                print("Installazione completata: riapri il programma.")
            raise SystemExit(0)

    messaggio = (f"Manca un componente necessario al programma: {errore}\n\n"
                 "Puoi installarlo aprendo il Prompt dei comandi e scrivendo:\n"
                 "    py -m pip install --user " + " ".join(PACCHETTI))
    print(messaggio)
    raise SystemExit(1)


try:
    _importa_librerie()
except ImportError as _errore:  # pragma: no cover
    _componenti_mancanti(_errore)

# ---------------------------------------------------------------------------
# Template della versione interattiva (HTML autonomo, nessuna dipendenza esterna)
# ---------------------------------------------------------------------------

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITOLO__</title>
<style>
  :root { --verde:#2e7d32; --rosso:#c62828; --bordo:#37474f; }
  * { box-sizing: border-box; }
  body { margin:0; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
         color:#212121; background:#fafafa; }
  header { padding:10px 16px; background:#fff; border-bottom:1px solid #e0e0e0; }
  header h1 { margin:0; font-size:17px; font-weight:600; }
  header p { margin:3px 0 0; font-size:12px; color:#616161; }
  #app { display:flex; height:calc(100vh - 56px); }
  #tela { flex:1; position:relative; overflow:hidden; background:#fff; }
  svg { width:100%; height:100%; cursor:grab; display:block; }
  svg.trascina { cursor:grabbing; }
  aside { width:290px; border-left:1px solid #e0e0e0; background:#fff; padding:12px 14px;
          overflow-y:auto; font-size:13px; }
  aside h2 { font-size:12px; text-transform:uppercase; letter-spacing:.06em; color:#757575;
             margin:14px 0 6px; font-weight:600; }
  aside h2:first-child { margin-top:0; }
  label.riga { display:flex; align-items:center; gap:7px; margin:4px 0; cursor:pointer; }
  input[type=search] { width:100%; padding:6px 8px; border:1px solid #cfd8dc; border-radius:5px; font-size:13px; }
  button { font-size:12px; padding:5px 9px; margin:3px 4px 3px 0; border:1px solid #cfd8dc;
           background:#fff; border-radius:5px; cursor:pointer; }
  button:hover { background:#eceff1; }
  table.dett { width:100%; border-collapse:collapse; font-size:12px; }
  table.dett td { padding:2px 0; border-bottom:1px solid #f0f0f0; }
  table.dett td:last-child { text-align:right; font-variant-numeric:tabular-nums; font-weight:600; }
  .pill { display:inline-block; padding:1px 7px; border-radius:9px; font-size:11px; border:1px solid #b0bec5; }
  .leg { display:flex; align-items:center; gap:7px; margin:3px 0; font-size:12px; }
  .pallino { width:13px; height:13px; border-radius:50%; border:1px solid var(--bordo); }
  .linea { width:22px; height:0; border-top-width:2px; }
  .nodo { cursor:pointer; }
  .nodo text { font-size:11px; font-weight:600; pointer-events:none;
               paint-order:stroke; stroke:#fff; stroke-width:3px;
               stroke-linejoin:round; }
  .off { display:none; }
  .spenta { opacity:.07; }
  .attenuato { opacity:.18; }
  footer { font-size:11px; color:#9e9e9e; margin-top:14px; line-height:1.45; }
</style>
</head>
<body>
<header>
  <h1>__TITOLO__</h1>
  <p>Trascina i nodi per riposizionarli · rotella per lo zoom · trascina lo sfondo per spostare la vista · clic su un nodo per isolarne le relazioni</p>
</header>

<div id="app">
  <div id="tela">
    <svg id="grafo" xmlns="http://www.w3.org/2000/svg">
      <defs>
        <marker id="fp" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7"
                markerUnits="userSpaceOnUse" orient="auto-start-reverse">
          <path d="M 0 0 L 10 5 L 0 10 z" fill="#2e7d32"></path>
        </marker>
        <marker id="fn" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7"
                markerUnits="userSpaceOnUse" orient="auto-start-reverse">
          <path d="M 0 0 L 10 5 L 0 10 z" fill="#c62828"></path>
        </marker>
      </defs>
      <g id="vista">
        <g id="archi"></g>
        <g id="nodi"></g>
      </g>
    </svg>
  </div>

  <aside>
    <h2>Livelli</h2>
    <label class="riga"><input type="checkbox" id="cbScelte" checked> Preferenze positive (verde)</label>
    <label class="riga"><input type="checkbox" id="cbRifiuti" checked> Rifiuti (rosso)</label>
    <label class="riga"><input type="checkbox" id="cbSoloRec"> Solo relazioni reciproche</label>
    <label class="riga"><input type="checkbox" id="cbEtichette" checked> Etichette</label>

    <h2>Vista</h2>
    <button id="btBersaglio">Layout a bersaglio</button>
    <button id="btForza">Riorganizza</button>
    <button id="btReset">Adatta</button>
    <button id="btSvg">Scarica SVG</button>

    <h2>Cerca</h2>
    <input type="search" id="ricerca" placeholder="ID o etichetta…" autocomplete="off">

    <h2>Dettaglio</h2>
    <div id="dettaglio"><p style="color:#9e9e9e;margin:0">Nessun nodo selezionato.</p></div>

    <h2>Legenda</h2>
    <div id="legenda"></div>
    <div class="leg"><span class="linea" style="border-top:2px solid var(--verde)"></span> vorrebbe lavorare con</div>
    <div class="leg"><span class="linea" style="border-top:4px solid var(--verde)"></span> preferenza reciproca</div>
    <div class="leg"><span class="linea" style="border-top:2px dashed var(--rosso)"></span> non vorrebbe lavorare con</div>
    <div class="leg"><span class="linea" style="border-top:4px dashed var(--rosso)"></span> rifiuto reciproco</div>
    <div class="leg"><span class="pallino" style="background:#fff;border:1px dashed #9e9e9e"></span> nessuna preferenza espressa</div>
    <footer>Dimensione del nodo = preferenze positive ricevute · spessore dell'anello esterno = indicazioni negative ricevute.
    Generato da sociogramma.py.</footer>
  </aside>
</div>

<script>
const DATI = __DATI__;
const COLORI = __COLORI__;

const svg = document.getElementById('grafo');
const gVista = document.getElementById('vista');
const gArchi = document.getElementById('archi');
const gNodi = document.getElementById('nodi');
const NS = 'http://www.w3.org/2000/svg';

const stato = { k: 1, tx: 0, ty: 0, sel: null };
const nodi = new Map(DATI.nodi.map(n => [n.id, Object.assign({}, n)]));
let posIniziale = new Map([...nodi].map(([k, n]) => [k, { x: n.x, y: n.y }]));

function raggio(n) { return 11 + 2.4 * n.scelte_ricevute; }
function bordo(n) { return 1 + Math.min(4, 0.9 * n.rifiuti_ricevuti); }

/* ---------- costruzione elementi ---------- */
const elemArco = [], elemNodo = new Map();

DATI.archi.forEach(a => {
  const p = document.createElementNS(NS, 'path');
  const positivo = a.tipo === 'pos';
  p.setAttribute('fill', 'none');
  p.setAttribute('stroke', positivo ? '#2e7d32' : '#c62828');
  p.setAttribute('stroke-width', a.rec ? 3.2 : 1.5);
  p.setAttribute('stroke-opacity', positivo ? 0.75 : 0.6);
  if (!positivo) p.setAttribute('stroke-dasharray', '6 4');
  p.setAttribute('marker-end', positivo ? 'url(#fp)' : 'url(#fn)');
  gArchi.appendChild(p);
  elemArco.push({ el: p, dati: a });
});

DATI.nodi.forEach(n => {
  const g = document.createElementNS(NS, 'g');
  g.setAttribute('class', 'nodo');
  const c = document.createElementNS(NS, 'circle');
  c.setAttribute('r', raggio(n));
  c.setAttribute('fill', COLORI[n.categoria] || '#eeeeee');
  c.setAttribute('stroke', n.compilazione === 'Assente' ? '#9e9e9e' : '#37474f');
  c.setAttribute('stroke-width', bordo(n));
  if (n.compilazione === 'Assente') c.setAttribute('stroke-dasharray', '4 3');
  const t = document.createElementNS(NS, 'text');
  t.setAttribute('text-anchor', 'middle');
  const parti = String(n.etichetta).split(' ');
  const linee = parti.length > 1 ? [parti[0], parti.slice(1).join(' ')] : [n.etichetta];
  linee.forEach((testo, i) => {
    const ts = document.createElementNS(NS, 'tspan');
    ts.setAttribute('x', 0);
    ts.setAttribute('y', raggio(n) + 16 + i * 12);
    ts.textContent = testo;
    t.appendChild(ts);
  });
  g.appendChild(c); g.appendChild(t);
  const tip = document.createElementNS(NS, 'title');
  tip.textContent = `${n.etichetta} — ${n.categoria}\npreferenze ricevute: ${n.scelte_ricevute} · rifiuti ricevuti: ${n.rifiuti_ricevuti}`;
  g.appendChild(tip);
  gNodi.appendChild(g);
  elemNodo.set(n.id, g);
  g.addEventListener('click', ev => { ev.stopPropagation(); seleziona(n.id); });
  attivaTrascinamento(g, n.id);
});

/* ---------- disegno ---------- */
function visibile(a) {
  if (a.tipo === 'pos' && !document.getElementById('cbScelte').checked) return false;
  if (a.tipo === 'neg' && !document.getElementById('cbRifiuti').checked) return false;
  if (document.getElementById('cbSoloRec').checked && !a.rec) return false;
  if (stato.sel && a.s !== stato.sel && a.t !== stato.sel) return false;
  return true;
}

function disegna() {
  const etichette = document.getElementById('cbEtichette').checked;
  elemArco.forEach(({ el, dati }) => {
    if (!visibile(dati)) { el.classList.add('off'); return; }
    el.classList.remove('off');
    const a = nodi.get(dati.s), b = nodi.get(dati.t);
    const dx = b.x - a.x, dy = b.y - a.y, d = Math.hypot(dx, dy) || 1;
    const curva = dati.tipo === 'pos' ? 0.13 : -0.17;
    const ux = dx / d, uy = dy / d;
    const x1 = a.x + ux * (raggio(a) + 2), y1 = a.y + uy * (raggio(a) + 2);
    const x2 = b.x - ux * (raggio(b) + 6), y2 = b.y - uy * (raggio(b) + 6);
    const mx = (x1 + x2) / 2 - (y2 - y1) * curva, my = (y1 + y2) / 2 + (x2 - x1) * curva;
    el.setAttribute('d', `M ${x1} ${y1} Q ${mx} ${my} ${x2} ${y2}`);
  });

  const vicini = new Set();
  if (stato.sel) {
    vicini.add(stato.sel);
    DATI.archi.forEach(a => { if (visibile(a)) { vicini.add(a.s); vicini.add(a.t); } });
  }
  nodi.forEach((n, id) => {
    const g = elemNodo.get(id);
    g.setAttribute('transform', `translate(${n.x},${n.y})`);
    g.querySelector('text').style.display = etichette ? '' : 'none';
    g.classList.toggle('attenuato', !!stato.sel && !vicini.has(id));
  });
}

function applicaVista() {
  gVista.setAttribute('transform', `translate(${stato.tx},${stato.ty}) scale(${stato.k})`);
}

/* ---------- interazione ---------- */
function coordinate(ev) {
  const r = svg.getBoundingClientRect();
  return { x: (ev.clientX - r.left - stato.tx) / stato.k, y: (ev.clientY - r.top - stato.ty) / stato.k };
}

function attivaTrascinamento(g, id) {
  let attivo = false;
  g.addEventListener('pointerdown', ev => {
    ev.stopPropagation(); attivo = true; g.setPointerCapture(ev.pointerId);
  });
  g.addEventListener('pointermove', ev => {
    if (!attivo) return;
    const p = coordinate(ev);
    const n = nodi.get(id); n.x = p.x; n.y = p.y;
    disegna();
  });
  const fine = ev => { attivo = false; try { g.releasePointerCapture(ev.pointerId); } catch (e) {} };
  g.addEventListener('pointerup', fine);
  g.addEventListener('pointercancel', fine);
}

let pan = null;
svg.addEventListener('pointerdown', ev => {
  pan = { x: ev.clientX, y: ev.clientY, tx: stato.tx, ty: stato.ty };
  svg.classList.add('trascina');
});
svg.addEventListener('pointermove', ev => {
  if (!pan) return;
  stato.tx = pan.tx + (ev.clientX - pan.x); stato.ty = pan.ty + (ev.clientY - pan.y);
  applicaVista();
});
['pointerup', 'pointerleave', 'pointercancel'].forEach(e =>
  svg.addEventListener(e, () => { pan = null; svg.classList.remove('trascina'); }));
svg.addEventListener('click', () => seleziona(null));

svg.addEventListener('wheel', ev => {
  ev.preventDefault();
  const r = svg.getBoundingClientRect();
  const px = ev.clientX - r.left, py = ev.clientY - r.top;
  const f = Math.exp(-ev.deltaY * 0.0015);
  const k = Math.max(0.15, Math.min(6, stato.k * f));
  stato.tx = px - (px - stato.tx) * (k / stato.k);
  stato.ty = py - (py - stato.ty) * (k / stato.k);
  stato.k = k; applicaVista();
}, { passive: false });

function seleziona(id) {
  stato.sel = (id === stato.sel) ? null : id;
  const box = document.getElementById('dettaglio');
  if (!stato.sel) {
    box.innerHTML = '<p style="color:#9e9e9e;margin:0">Nessun nodo selezionato.</p>';
  } else {
    const n = nodi.get(stato.sel);
    const scelti = DATI.archi.filter(a => a.s === n.id && a.tipo === 'pos').map(a => nodi.get(a.t).etichetta);
    const rifiutati = DATI.archi.filter(a => a.s === n.id && a.tipo === 'neg').map(a => nodi.get(a.t).etichetta);
    const da = DATI.archi.filter(a => a.t === n.id && a.tipo === 'pos').map(a => nodi.get(a.s).etichetta);
    const contro = DATI.archi.filter(a => a.t === n.id && a.tipo === 'neg').map(a => nodi.get(a.s).etichetta);
    const rec = DATI.archi.filter(a => a.s === n.id && a.tipo === 'pos' && a.rec).map(a => nodi.get(a.t).etichetta);
    box.innerHTML =
      `<p style="margin:0 0 6px"><strong>${n.etichetta}</strong> <span class="pill"
        style="background:${COLORI[n.categoria] || '#eee'}">${n.categoria}</span></p>
       <table class="dett">
         <tr><td>Preferenze ricevute</td><td>${n.scelte_ricevute}</td></tr>
         <tr><td>Indicazioni negative</td><td>${n.rifiuti_ricevuti}</td></tr>
         <tr><td>Status sociometrico</td><td>${n.status}</td></tr>
         <tr><td>Preferenza sociale (z)</td><td>${n.preferenza}</td></tr>
         <tr><td>Impatto sociale (z)</td><td>${n.impatto}</td></tr>
         <tr><td>Compilazione</td><td>${n.compilazione}</td></tr>
       </table>
       <p style="margin:8px 0 2px"><em>Vorrebbe lavorare con:</em> ${scelti.join(', ') || '—'}</p>
       <p style="margin:2px 0"><em>Scelto da:</em> ${da.join(', ') || '—'}</p>
       <p style="margin:2px 0"><em>Preferenze reciproche:</em> ${rec.join(', ') || '—'}</p>
       <p style="margin:2px 0"><em>Non vorrebbe lavorare con:</em> ${rifiutati.join(', ') || '—'}</p>
       <p style="margin:2px 0"><em>Evitato da:</em> ${contro.join(', ') || '—'}</p>`;
  }
  disegna();
}

['cbScelte', 'cbRifiuti', 'cbSoloRec', 'cbEtichette'].forEach(id =>
  document.getElementById(id).addEventListener('change', disegna));

document.getElementById('ricerca').addEventListener('input', ev => {
  const q = ev.target.value.trim().toLowerCase();
  if (!q) { seleziona(null); return; }
  const trovato = [...nodi.values()].find(n =>
    n.etichetta.toLowerCase().includes(q) || String(n.id).toLowerCase().includes(q));
  if (trovato) { stato.sel = trovato.id; seleziona(trovato.id); }
});

document.getElementById('btBersaglio').addEventListener('click', () => {
  posIniziale.forEach((p, id) => { const n = nodi.get(id); n.x = p.x; n.y = p.y; });
  disegna(); adatta();
});

document.getElementById('btForza').addEventListener('click', () => forza(260));
document.getElementById('btReset').addEventListener('click', adatta);

document.getElementById('btSvg').addEventListener('click', () => {
  const copia = svg.cloneNode(true);
  copia.setAttribute('xmlns', NS);
  copia.querySelectorAll('.off').forEach(e => e.remove());
  const b = new Blob(['<?xml version="1.0" encoding="UTF-8"?>\n' + copia.outerHTML],
                     { type: 'image/svg+xml' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(b); a.download = 'sociogramma.svg'; a.click();
  URL.revokeObjectURL(a.href);
});

/* ---------- simulazione a forze (solo scelte positive) ---------- */
function forza(passi) {
  const lista = [...nodi.values()];
  const archi = DATI.archi.filter(a => a.tipo === 'pos');
  const cx = lista.reduce((s, n) => s + n.x, 0) / lista.length;
  const cy = lista.reduce((s, n) => s + n.y, 0) / lista.length;
  let i = 0;
  (function passo() {
    const fx = new Map(), fy = new Map();
    lista.forEach(n => { fx.set(n.id, 0); fy.set(n.id, 0); });
    for (let a = 0; a < lista.length; a++) {
      for (let b = a + 1; b < lista.length; b++) {
        const A = lista[a], B = lista[b];
        let dx = A.x - B.x, dy = A.y - B.y;
        let d2 = dx * dx + dy * dy || 1;
        const rep = 90000 / d2;
        const d = Math.sqrt(d2);
        fx.set(A.id, fx.get(A.id) + rep * dx / d); fy.set(A.id, fy.get(A.id) + rep * dy / d);
        fx.set(B.id, fx.get(B.id) - rep * dx / d); fy.set(B.id, fy.get(B.id) - rep * dy / d);
      }
    }
    archi.forEach(e => {
      const A = nodi.get(e.s), B = nodi.get(e.t);
      const dx = B.x - A.x, dy = B.y - A.y, d = Math.hypot(dx, dy) || 1;
      const k = 0.012 * (d - 150) * (e.rec ? 1.8 : 1);
      fx.set(A.id, fx.get(A.id) + k * dx / d); fy.set(A.id, fy.get(A.id) + k * dy / d);
      fx.set(B.id, fx.get(B.id) - k * dx / d); fy.set(B.id, fy.get(B.id) - k * dy / d);
    });
    const raffredda = 1 - i / passi;
    lista.forEach(n => {
      n.x += Math.max(-25, Math.min(25, fx.get(n.id))) * raffredda * 0.5 + (cx - n.x) * 0.004;
      n.y += Math.max(-25, Math.min(25, fy.get(n.id))) * raffredda * 0.5 + (cy - n.y) * 0.004;
    });
    disegna();
    if (++i < passi) requestAnimationFrame(passo); else adatta();
  })();
}

/* ---------- adattamento vista ---------- */
function adatta() {
  const lista = [...nodi.values()];
  const minx = Math.min(...lista.map(n => n.x - raggio(n))) - 30;
  const maxx = Math.max(...lista.map(n => n.x + raggio(n))) + 30;
  const miny = Math.min(...lista.map(n => n.y - raggio(n))) - 30;
  const maxy = Math.max(...lista.map(n => n.y + raggio(n))) + 30;
  const r = svg.getBoundingClientRect();
  const k = Math.min(r.width / (maxx - minx), r.height / (maxy - miny));
  stato.k = Math.max(0.15, Math.min(3, k));
  stato.tx = (r.width - (maxx - minx) * stato.k) / 2 - minx * stato.k;
  stato.ty = (r.height - (maxy - miny) * stato.k) / 2 - miny * stato.k;
  applicaVista();
}

const legenda = document.getElementById('legenda');
Object.entries(COLORI).forEach(([k, c]) => {
  const d = document.createElement('div');
  d.className = 'leg';
  d.innerHTML = `<span class="pallino" style="background:${c}"></span> ${k}`;
  legenda.appendChild(d);
});

disegna();
window.addEventListener('resize', adatta);
adatta();
</script>
</body>
</html>
"""

# ---------------------------------------------------------------------------
# Configurazione
# ---------------------------------------------------------------------------

COL_ID = "ID"
COL_SI = ["SI-1", "SI-2", "SI-3"]
COL_NO = ["NO-1", "NO-2"]
COL_COGNOME = "COGNOME"
COL_NOME = "NOME"

# Pesi delle nomine ordinate (la prima scelta pesa più della terza)
PESI_SI = {"SI-1": 3.0, "SI-2": 2.0, "SI-3": 1.0}
PESI_NO = {"NO-1": 2.0, "NO-2": 1.0}

COLORI_CATEGORIA = {
    "Popolare":    "#a5d6a7",
    "Rifiutato":   "#ef9a9a",
    "Controverso": "#ffe082",
    "Trascurato":  "#e0e0e0",
    "Medio":       "#90caf9",
}
VERDE = "#2e7d32"
ROSSO = "#c62828"


# ---------------------------------------------------------------------------
# Lettura e normalizzazione dei dati
# ---------------------------------------------------------------------------

def _chiave(nome) -> str:
    """Normalizza un nome di colonna: 'si 1' -> 'SI1', 'NO_2' -> 'NO2'."""
    return "".join(str(nome).upper().split()).replace("-", "").replace("_", "")


def _as_id(valore):
    """Converte una cella in un ID stringa pulito, oppure None se vuota."""
    if valore is None:
        return None
    if isinstance(valore, float):
        if math.isnan(valore):
            return None
        if valore.is_integer():
            return str(int(valore))
        return str(valore)
    try:
        if pd.isna(valore):
            return None
    except (TypeError, ValueError):
        pass
    testo = str(valore).strip()
    if testo == "" or testo.lower() in {"nan", "none", "<na>", "-", "/"}:
        return None
    return testo


def _due_righe(testo: str, soglia: int = 11) -> str:
    """Manda a capo 'Mario Rossi' quando l'etichetta è lunga."""
    if len(testo) <= soglia or " " not in testo:
        return testo
    nome, resto = testo.split(" ", 1)
    return f"{nome}\n{resto}"


def carica_dati(percorso: Path, foglio, col_etichetta: str | None = None):
    """Legge l'Excel e restituisce (DataFrame canonico, etichette, anagrafica)."""
    df = pd.read_excel(percorso, sheet_name=foglio)
    if isinstance(df, dict):
        df = next(iter(df.values()))

    mappa = {_chiave(c): c for c in df.columns}
    attese = [COL_ID] + COL_SI + COL_NO
    mancanti = [c for c in attese if _chiave(c) not in mappa]
    if mancanti:
        raise SystemExit(
            "Colonne mancanti nel file: " + ", ".join(mancanti) +
            "\nColonne trovate: " + ", ".join(str(c) for c in df.columns))

    out = pd.DataFrame(index=df.index)
    for c in attese:
        # dtype=object: evita che i valori mancanti tornino come NaN invece di None
        out[c] = pd.Series([_as_id(v) for v in df[mappa[_chiave(c)]]],
                           index=df.index, dtype=object)

    etichette, anagrafica = {}, {}
    k_cognome = _chiave(COL_COGNOME) if _chiave(COL_COGNOME) in mappa else None
    k_nome = _chiave(COL_NOME) if _chiave(COL_NOME) in mappa else None
    k_libera = None
    if col_etichetta:
        k_libera = _chiave(col_etichetta)
        if k_libera not in mappa:
            raise SystemExit(f"Colonna etichetta non trovata: {col_etichetta}")

    for posizione, ident in enumerate(out[COL_ID]):
        if not ident:
            continue
        cognome = _as_id(df[mappa[k_cognome]].iloc[posizione]) if k_cognome else None
        nome = _as_id(df[mappa[k_nome]].iloc[posizione]) if k_nome else None
        anagrafica[ident] = {"Nome": nome or "", "Cognome": cognome or ""}
        if k_libera:
            etichetta = _as_id(df[mappa[k_libera]].iloc[posizione])
        else:
            etichetta = " ".join(x for x in (nome, cognome) if x)
        etichette[ident] = etichetta or ident

    # omonimi: si aggiunge l'ID per distinguerli
    quante = {}
    for valore in etichette.values():
        quante[valore] = quante.get(valore, 0) + 1
    for ident, valore in list(etichette.items()):
        if quante.get(valore, 0) > 1:
            etichette[ident] = f"{valore} ({ident})"

    out = out[out[COL_ID].notna()].reset_index(drop=True)
    if out.empty:
        raise SystemExit("Nessuna riga valida: la colonna ID risulta vuota.")
    return out, etichette, anagrafica


# ---------------------------------------------------------------------------
# Costruzione del grafo
# ---------------------------------------------------------------------------

def costruisci_grafo(df: pd.DataFrame, includi_esterni: bool = False):
    """Ritorna (grafo diretto, lista anomalie)."""
    soggetti = list(dict.fromkeys(df[COL_ID].tolist()))
    anomalie = []

    for duplicato in df[COL_ID][df[COL_ID].duplicated()].unique().tolist():
        anomalie.append({"ID": duplicato, "Tipo": "ID duplicato",
                         "Dettaglio": "più righe con lo stesso ID"})

    G = nx.DiGraph()
    G.add_nodes_from(soggetti)

    for _, riga in df.iterrows():
        sorgente = _as_id(riga[COL_ID])
        for colonna in COL_SI + COL_NO:
            destinazione = _as_id(riga[colonna])
            if destinazione is None:
                continue
            positivo = colonna in PESI_SI
            peso = PESI_SI[colonna] if positivo else PESI_NO[colonna]

            if destinazione == sorgente:
                anomalie.append({"ID": sorgente, "Tipo": "Auto-nomina",
                                 "Dettaglio": f"{colonna} punta a se stesso"})
                continue
            if destinazione not in soggetti:
                anomalie.append({"ID": sorgente, "Tipo": "ID non trovato",
                                 "Dettaglio": f"{colonna} = {destinazione}"})
                if not includi_esterni:
                    continue
                G.add_node(destinazione, esterno=True)
            if G.has_edge(sorgente, destinazione):
                anomalie.append({"ID": sorgente, "Tipo": "Nomina ripetuta",
                                 "Dettaglio": f"{destinazione} indicato più volte"})
                continue

            G.add_edge(sorgente, destinazione, positivo=positivo, peso=peso, rango=colonna)

    for u, v, dati in G.edges(data=True):
        contrario = G.get_edge_data(v, u)
        dati["reciproca"] = bool(contrario and contrario["positivo"] == dati["positivo"])

    return G, anomalie


# ---------------------------------------------------------------------------
# Indici sociometrici
# ---------------------------------------------------------------------------

def conteggio_compilazione(df: pd.DataFrame) -> dict:
    """Per ogni persona quante caselle SI e NO ha effettivamente compilato."""
    conteggio = {}
    for _, riga in df.iterrows():
        ident = _as_id(riga[COL_ID])
        if not ident:
            continue
        conteggio[ident] = {
            "SI": sum(1 for c in COL_SI if _as_id(riga[c])),
            "NO": sum(1 for c in COL_NO if _as_id(riga[c])),
        }
    return conteggio


def _z(serie: pd.Series) -> pd.Series:
    sd = serie.std(ddof=0)
    if not sd or np.isnan(sd):
        return pd.Series(0.0, index=serie.index)
    return (serie - serie.mean()) / sd


def calcola_indici(G: nx.DiGraph, etichette: dict, anagrafica: dict | None = None,
                   compilazione: dict | None = None) -> pd.DataFrame:
    nodi = list(G.nodes())
    n = len(nodi)
    compilazione = compilazione or {}
    # chi non ha compilato nulla non puo' scegliere nessuno: non entra nel
    # denominatore degli indici degli altri
    partecipanti = {nodo for nodo in nodi
                    if sum(compilazione.get(nodo, {"SI": 1, "NO": 1}).values()) > 0}
    righe = []
    for nodo in nodi:
        in_pos = [(u, d) for u, _, d in G.in_edges(nodo, data=True) if d["positivo"]]
        in_neg = [(u, d) for u, _, d in G.in_edges(nodo, data=True) if not d["positivo"]]
        out_pos = [d for _, _, d in G.out_edges(nodo, data=True) if d["positivo"]]
        out_neg = [d for _, _, d in G.out_edges(nodo, data=True) if not d["positivo"]]
        righe.append({
            "ID": nodo,
            "Etichetta": etichette.get(nodo, nodo),
            "Cognome": (anagrafica or {}).get(nodo, {}).get("Cognome", ""),
            "Nome": (anagrafica or {}).get(nodo, {}).get("Nome", ""),
            "Scelte ricevute": len(in_pos),
            "Rifiuti ricevuti": len(in_neg),
            "Scelte ricevute (pesate)": sum(d["peso"] for _, d in in_pos),
            "Rifiuti ricevuti (pesati)": sum(d["peso"] for _, d in in_neg),
            "Scelte emesse": len(out_pos),
            "Rifiuti emessi": len(out_neg),
            "Scelte reciproche": sum(1 for _, d in in_pos if d["reciproca"]),
            "Rifiuti reciproci": sum(1 for _, d in in_neg if d["reciproca"]),
            "SI compilate": compilazione.get(nodo, {}).get("SI", len(out_pos)),
            "NO compilate": compilazione.get(nodo, {}).get("NO", len(out_neg)),
            "Potenziali sceglienti": max(len(partecipanti) - (1 if nodo in partecipanti else 0), 1),
        })

    df = pd.DataFrame(righe)
    denom = df["Potenziali sceglienti"]
    df["Status sociometrico"] = ((df["Scelte ricevute"] - df["Rifiuti ricevuti"]) / denom).round(3)
    df["Indice di popolarità"] = (df["Scelte ricevute"] / denom).round(3)
    df["Indice di rifiuto"] = (df["Rifiuti ricevuti"] / denom).round(3)
    df["Compilazione"] = np.where(
        df["SI compilate"] + df["NO compilate"] == 0, "Assente",
        np.where((df["SI compilate"] == len(COL_SI)) & (df["NO compilate"] == len(COL_NO)),
                 "Completa", "Parziale"))

    zp, zn = _z(df["Scelte ricevute"]), _z(df["Rifiuti ricevuti"])
    df["Preferenza sociale (z)"] = (zp - zn).round(3)
    df["Impatto sociale (z)"] = (zp + zn).round(3)

    def categoria(i):
        sp = df.loc[i, "Preferenza sociale (z)"]
        si = df.loc[i, "Impatto sociale (z)"]
        p, r = zp[i], zn[i]
        if sp > 1.0 and p > 0 and r < 0:
            return "Popolare"
        if sp < -1.0 and p < 0 and r > 0:
            return "Rifiutato"
        if si > 1.0 and p > 0 and r > 0:
            return "Controverso"
        if si < -1.0 and p <= 0 and r <= 0:
            return "Trascurato"
        return "Medio"

    df["Categoria"] = [categoria(i) for i in df.index]
    df["Isolato"] = (df["Scelte ricevute"] == 0) & (df["Rifiuti ricevuti"] == 0)
    return (df.sort_values(["Status sociometrico", "Scelte ricevute"], ascending=False)
              .reset_index(drop=True))


def coppie_reciproche(G: nx.DiGraph, positive: bool, etichette: dict) -> pd.DataFrame:
    viste, righe = set(), []
    for u, v, d in G.edges(data=True):
        if d["positivo"] is positive and d["reciproca"] and (v, u) not in viste:
            viste.add((u, v))
            righe.append({"A": etichette.get(u, u), "B": etichette.get(v, v),
                          "Peso A->B": d["peso"], "Peso B->A": G[v][u]["peso"]})
    return pd.DataFrame(righe, columns=["A", "B", "Peso A->B", "Peso B->A"])


def sottogruppi(G: nx.DiGraph):
    """Componenti e cricche del grafo delle sole scelte reciproche."""
    R = nx.Graph()
    R.add_nodes_from(G.nodes())
    R.add_edges_from([(u, v) for u, v, d in G.edges(data=True)
                      if d["positivo"] and d["reciproca"]])
    componenti = [sorted(c) for c in nx.connected_components(R) if len(c) >= 2]
    cricche = [sorted(c) for c in nx.find_cliques(R) if len(c) >= 3]
    return sorted(componenti, key=len, reverse=True), sorted(cricche, key=len, reverse=True)


def indici_di_gruppo(G: nx.DiGraph, indici: pd.DataFrame) -> pd.DataFrame:
    n = G.number_of_nodes()
    assenti = int((indici["Compilazione"] == "Assente").sum())
    parziali = int((indici["Compilazione"] == "Parziale").sum())
    coppie_possibili = n * (n - 1) / 2 if n > 1 else 1
    archi_pos = [d for _, _, d in G.edges(data=True) if d["positivo"]]
    archi_neg = [d for _, _, d in G.edges(data=True) if not d["positivo"]]
    rec_pos = sum(1 for d in archi_pos if d["reciproca"]) / 2
    rec_neg = sum(1 for d in archi_neg if d["reciproca"]) / 2

    voci = [
        ("Persone", n),
        ("Hanno compilato il questionario", n - assenti),
        ("Compilazioni parziali", parziali),
        ("Nessuna preferenza espressa", assenti),
        ("Caselle SI compilate su quelle possibili",
         f"{int(indici['SI compilate'].sum())}/{n * len(COL_SI)}"),
        ("Caselle NO compilate su quelle possibili",
         f"{int(indici['NO compilate'].sum())}/{n * len(COL_NO)}"),
        ("Scelte espresse", len(archi_pos)),
        ("Rifiuti espressi", len(archi_neg)),
        ("Scelte reciproche (coppie)", int(rec_pos)),
        ("Rifiuti reciproci (coppie)", int(rec_neg)),
        ("Densità delle scelte", round(len(archi_pos) / (n * (n - 1)), 3) if n > 1 else 0.0),
        ("Indice di coesione", round(rec_pos / coppie_possibili, 3)),
        ("Indice di dissociazione", round(rec_neg / coppie_possibili, 3)),
        ("Indice di reciprocità delle scelte",
         round(2 * rec_pos / len(archi_pos), 3) if archi_pos else 0.0),
        ("Persone senza nomine ricevute", int(indici["Isolato"].sum())),
        ("% senza nomine ricevute", round(100 * indici["Isolato"].mean(), 1)),
    ]
    for categoria in COLORI_CATEGORIA:
        voci.append((f"Persone '{categoria}'", int((indici["Categoria"] == categoria).sum())))
    return pd.DataFrame(voci, columns=["Indicatore", "Valore"], dtype=object)


def matrice_sociometrica(G: nx.DiGraph, indici: pd.DataFrame) -> pd.DataFrame:
    ordine = indici["ID"].tolist()
    etichette = dict(zip(indici["ID"], indici["Etichetta"]))
    M = pd.DataFrame(0, index=[etichette[i] for i in ordine],
                     columns=[etichette[i] for i in ordine], dtype=int)
    for u, v, d in G.edges(data=True):
        if u in etichette and v in etichette:
            M.loc[etichette[u], etichette[v]] = int(d["peso"]) * (1 if d["positivo"] else -1)
    M.index.name = "sceglie \\ scelto"
    return M


# ---------------------------------------------------------------------------
# Layout e grafici
# ---------------------------------------------------------------------------

def calcola_layout(G: nx.DiGraph, indici: pd.DataFrame, tipo: str, seed: int):
    if tipo == "bersaglio":
        ordinati = indici.sort_values("Status sociometrico", ascending=False)["ID"].tolist()
        n = len(ordinati)
        n_anelli = min(4, max(1, math.ceil(n / 7)))
        per_anello = math.ceil(n / n_anelli)
        base = max(0.26, 0.075 * per_anello)
        passo = max(0.32, base * 0.85)
        raggi, pos = [], {}
        for anello in range(n_anelli):
            gruppo = ordinati[anello * per_anello:(anello + 1) * per_anello]
            if not gruppo:
                continue
            raggio = 0.0 if (anello == 0 and len(gruppo) == 1) else base + passo * anello
            sfasamento = (math.pi / len(gruppo)) * (anello % 2)
            raggi.append(raggio)
            for k, nodo in enumerate(gruppo):
                angolo = 2 * math.pi * k / len(gruppo) + sfasamento
                pos[nodo] = np.array([raggio * math.cos(angolo), raggio * math.sin(angolo)])
        if n == 1:
            pos[ordinati[0]] = np.array([0.0, 0.0])
        return pos, raggi

    if tipo == "circolare":
        return nx.circular_layout(G), []

    grafo_pos = nx.DiGraph()
    grafo_pos.add_nodes_from(G.nodes())
    grafo_pos.add_edges_from([(u, v) for u, v, d in G.edges(data=True) if d["positivo"]])
    if tipo == "kamada":
        base = grafo_pos.to_undirected()
        pos = (nx.kamada_kawai_layout(base) if base.number_of_edges()
               else nx.circular_layout(base))
    else:
        pos = nx.spring_layout(grafo_pos, seed=seed,
                               k=1.5 / math.sqrt(max(len(G), 1)), iterations=500)
    return pos, []


def disegna(G, indici, pos, anelli, percorso: Path, titolo: str,
            mostra_positivi=True, mostra_negativi=True, dpi=200, mostra_anelli=True,
            etichette_sotto=None):
    info = indici.set_index("ID")
    nodi = list(G.nodes())
    metrica = "Scelte ricevute" if mostra_positivi else "Rifiuti ricevuti"

    def valore(nodo, colonna):
        return int(info.loc[nodo, colonna]) if nodo in info.index else 0

    dimensioni = [420 + 280 * valore(n, metrica) for n in nodi]
    colori = [COLORI_CATEGORIA.get(info.loc[n, "Categoria"], "#eeeeee")
              if n in info.index else "#ffffff" for n in nodi]
    etichette = {n: (info.loc[n, "Etichetta"] if n in info.index else n) for n in nodi}

    lato = max(10.0, 1.2 * math.sqrt(max(len(nodi), 1)) + 6.0)
    fig, ax = plt.subplots(figsize=(lato, lato))
    ax.set_axis_off()
    ax.set_aspect("equal")

    if anelli and mostra_anelli:
        bordi = [(a + b) / 2 for a, b in zip(anelli, anelli[1:])]
        ultimo = (anelli[-1] - anelli[-2]) if len(anelli) > 1 else 0.3
        bordi.append(anelli[-1] + ultimo / 2)
        for raggio in bordi:
            ax.add_patch(Circle((0, 0), raggio, fill=False, ls=":", lw=0.8,
                                ec="#bdbdbd", zorder=0))

    def archi(positivo, reciproci):
        return [(u, v) for u, v, d in G.edges(data=True)
                if d["positivo"] is positivo and d["reciproca"] is reciproci]

    gruppi = []
    if mostra_positivi:
        gruppi.append((archi(True, False), VERDE, "solid", 1.5, 0.75, 0.12))
        gruppi.append((archi(True, True), VERDE, "solid", 3.2, 0.95, 0.06))
    if mostra_negativi:
        gruppi.append((archi(False, False), ROSSO, "dashed", 1.3, 0.6, -0.16))
        gruppi.append((archi(False, True), ROSSO, "dashed", 2.8, 0.85, -0.08))

    for lista, colore, stile, spessore, alpha, curva in gruppi:
        if not lista:
            continue
        nx.draw_networkx_edges(
            G, pos, edgelist=lista, ax=ax, edge_color=colore, style=stile,
            width=spessore, alpha=alpha, arrows=True, arrowsize=15, arrowstyle="-|>",
            node_size=dimensioni, connectionstyle=f"arc3,rad={curva}",
            min_source_margin=6, min_target_margin=10)

    assenti = {n for n in nodi if n in info.index and info.loc[n, "Compilazione"] == "Assente"}
    for elenco, tratteggio in ((sorted(set(nodi) - assenti, key=nodi.index), False),
                               (sorted(assenti, key=nodi.index), True)):
        if not elenco:
            continue
        indici_nodo = [nodi.index(x) for x in elenco]
        collezione = nx.draw_networkx_nodes(
            G, pos, nodelist=elenco, ax=ax,
            node_size=[dimensioni[i] for i in indici_nodo],
            node_color=[colori[i] for i in indici_nodo],
            edgecolors="#9e9e9e" if tratteggio else "#37474f",
            linewidths=1.4 if tratteggio else 1.2)
        if tratteggio:
            collezione.set_linestyle((0, (3, 2)))
    lunghezza = max((len(v) for v in etichette.values()), default=0)
    sotto = lunghezza > 7 if etichette_sotto is None else etichette_sotto
    if sotto:
        # etichette sotto al nodo: i nomi completi non entrerebbero nel pallino
        ax.margins(0.06)
        fig.canvas.draw()
        trasforma = ax.transData.transform
        inversa = ax.transData.inverted().transform
        for nodo, dimensione in zip(nodi, dimensioni):
            x, y = pos[nodo]
            raggio_punti = math.sqrt(dimensione / math.pi)
            px, py = trasforma((x, y))
            _, y_testo = inversa((px, py - (raggio_punti * fig.dpi / 72 + 3)))
            ax.text(x, y_testo, _due_righe(etichette[nodo]), ha="center", va="top",
                    fontsize=8.5, fontweight="bold", color="#212121", zorder=5,
                    linespacing=1.05,
                    bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none", alpha=0.7))
    else:
        nx.draw_networkx_labels(G, pos, labels=etichette, ax=ax, font_size=9,
                                font_weight="bold", font_color="#212121")

    legenda = [Line2D([], [], marker="o", ls="", markersize=11, markeredgecolor="#37474f",
                      markerfacecolor=c, label=k) for k, c in COLORI_CATEGORIA.items()]
    if mostra_positivi:
        legenda += [Line2D([], [], color=VERDE, lw=1.6, label="Preferenza"),
                    Line2D([], [], color=VERDE, lw=3.2, label="Preferenza reciproca")]
    if mostra_negativi:
        legenda += [Line2D([], [], color=ROSSO, lw=1.4, ls="--", label="Rifiuto"),
                    Line2D([], [], color=ROSSO, lw=2.8, ls="--", label="Rifiuto reciproco")]
    if assenti:
        legenda.append(Line2D([], [], marker="o", ls="", markersize=11, markerfacecolor="white",
                              markeredgecolor="#9e9e9e",
                              label="nessuna preferenza espressa\n(bordo tratteggiato)"))
    ax.legend(handles=legenda, loc="upper left", bbox_to_anchor=(1.0, 1.0),
              frameon=False, fontsize=9)

    ax.set_title(titolo, fontsize=15, pad=18)
    ax.margins(0.06)
    fig.tight_layout()
    fig.savefig(percorso, dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return percorso


# ---------------------------------------------------------------------------
# Versione interattiva (HTML autonomo, funziona offline)
# ---------------------------------------------------------------------------

def scrivi_html(G, indici, pos, percorso: Path, titolo: str) -> Path:
    info = indici.set_index("ID")
    xs = [p[0] for p in pos.values()] or [0]
    ys = [p[1] for p in pos.values()] or [0]
    minx, maxx, miny, maxy = min(xs), max(xs), min(ys), max(ys)
    amp = max(maxx - minx, maxy - miny) or 1.0
    scala = 820 / amp

    nodi = []
    for nodo in G.nodes():
        x, y = pos[nodo]
        r = info.loc[nodo] if nodo in info.index else None
        nodi.append({
            "id": str(nodo),
            "etichetta": str(r["Etichetta"]) if r is not None else str(nodo),
            "categoria": str(r["Categoria"]) if r is not None else "Medio",
            "scelte_ricevute": int(r["Scelte ricevute"]) if r is not None else 0,
            "rifiuti_ricevuti": int(r["Rifiuti ricevuti"]) if r is not None else 0,
            "status": float(r["Status sociometrico"]) if r is not None else 0.0,
            "preferenza": float(r["Preferenza sociale (z)"]) if r is not None else 0.0,
            "impatto": float(r["Impatto sociale (z)"]) if r is not None else 0.0,
            "compilazione": str(r["Compilazione"]) if r is not None else "Assente",
            "x": round((x - (minx + maxx) / 2) * scala + 450, 1),
            "y": round(((miny + maxy) / 2 - y) * scala + 420, 1),
        })
    archi = [{"s": str(u), "t": str(v), "tipo": "pos" if d["positivo"] else "neg",
              "peso": d["peso"], "rec": bool(d["reciproca"])}
             for u, v, d in G.edges(data=True)]

    html = (HTML_TEMPLATE
            .replace("__TITOLO__", titolo)
            .replace("__DATI__", json.dumps({"nodi": nodi, "archi": archi}, ensure_ascii=False))
            .replace("__COLORI__", json.dumps(COLORI_CATEGORIA, ensure_ascii=False)))
    percorso.write_text(html, encoding="utf-8")
    return percorso


# ---------------------------------------------------------------------------
# Report commentato (.docx)
# ---------------------------------------------------------------------------

def _elenco(indici: pd.DataFrame, categoria: str) -> list[str]:
    return indici.loc[indici["Categoria"] == categoria, "Etichetta"].tolist()


def _giudizio(valore: float, soglie=(0.10, 0.25)) -> str:
    if valore < soglie[0]:
        return "basso"
    if valore < soglie[1]:
        return "medio"
    return "alto"


def commento_automatico(indici: pd.DataFrame, gruppo: pd.DataFrame,
                        componenti, cricche) -> list[str]:
    g = dict(zip(gruppo["Indicatore"], gruppo["Valore"]))
    n = int(g["Persone"])
    testo = []

    testo.append(
        f"Il gruppo è composto da {n} persone, che hanno espresso complessivamente "
        f"{int(g['Scelte espresse'])} preferenze positive e {int(g['Rifiuti espressi'])} negative. "
        f"La densità delle scelte è {g['Densità delle scelte']}, cioè la quota di "
        "legami positivi effettivamente attivati rispetto a tutti quelli possibili.")

    assenti = indici.loc[indici["Compilazione"] == "Assente", "Etichetta"].tolist()
    parziali = indici.loc[indici["Compilazione"] == "Parziale", "Etichetta"].tolist()
    if assenti or parziali:
        frase = ""
        if assenti:
            frase += (f"{len(assenti)} persone non hanno espresso alcuna preferenza "
                      f"({', '.join(assenti)}): compaiono nei diagrammi con il bordo tratteggiato "
                      "e possono solo ricevere indicazioni, non darne. Gli indici degli altri sono "
                      "calcolati escludendole dal numero dei potenziali sceglienti. ")
        if parziali:
            frase += (f"Hanno compilato solo in parte il questionario {', '.join(parziali)}: "
                      "il numero di preferenze ricevute dagli altri risulta di conseguenza più basso.")
        testo.append(frase.strip())
        if len(assenti) > 0.2 * n:
            testo.append("La quota di mancate risposte è rilevante: gli indici vanno letti con "
                         "prudenza, perché una parte della rete non è osservabile.")

    coesione = float(g["Indice di coesione"])
    testo.append(
        f"L'indice di coesione vale {coesione} ({_giudizio(coesione)}): misura la quota di "
        f"coppie che si sono scelte reciprocamente ({int(g['Scelte reciproche (coppie)'])} coppie). "
        f"L'indice di dissociazione, basato sui rifiuti reciproci, vale "
        f"{g['Indice di dissociazione']}.")

    popolari = _elenco(indici, "Popolare")
    if popolari:
        top = indici.nlargest(3, "Scelte ricevute")[["Etichetta", "Scelte ricevute"]]
        elenco = ", ".join(f"{r.Etichetta} ({int(r._2)} scelte)" for r in top.itertuples())
        testo.append(
            f"Emergono {len(popolari)} persone con status elevato ({', '.join(popolari)}). "
            f"I più richiesti in assoluto sono {elenco}: sono i riferimenti naturali del gruppo e "
            "possono fare da capofila o da punto di aggregazione nella composizione dei team.")
    else:
        testo.append("Nessuno raggiunge una preferenza sociale nettamente superiore alla media: "
                     "il gruppo non esprime riferimenti informali marcati.")

    rifiutati = _elenco(indici, "Rifiutato")
    if rifiutati:
        testo.append(
            f"Risultano poco desiderati come compagni di lavoro {len(rifiutati)} persone "
            f"({', '.join(rifiutati)}): ricevono molte indicazioni negative e poche positive. Sono "
            "le situazioni che meritano più attenzione, verificando con l'osservazione diretta se "
            "dipendono da conflitti, da ruoli scomodi o da difficoltà relazionali.")

    controversi = _elenco(indici, "Controverso")
    if controversi:
        testo.append(
            f"Posizione controversa per {', '.join(controversi)}: ricevono insieme molte preferenze "
            "positive e molte negative, segno di forte visibilità e di un gruppo polarizzato nei "
            "loro confronti.")

    trascurati = _elenco(indici, "Trascurato")
    isolati = indici.loc[indici["Isolato"], "Etichetta"].tolist()
    if trascurati or isolati:
        frase = ""
        if trascurati:
            frase += (f"Risultano poco considerati {', '.join(trascurati)}: ricevono poche "
                      "indicazioni di qualsiasi segno, quindi passano inosservati più che essere "
                      "respinti. ")
        if isolati:
            frase += (f"In particolare {', '.join(isolati)} non compaiono in nessuna indicazione "
                      "ricevuta: nessuno li ha nominati, né in positivo né in negativo.")
        testo.append(frase.strip())

    if componenti:
        descrizione = "; ".join(" - ".join(c) for c in componenti[:6])
        testo.append(
            f"La rete delle preferenze reciproche si articola in {len(componenti)} aggregati: {descrizione}"
            + (" (elenco troncato)." if len(componenti) > 6 else "."))
    if cricche:
        descrizione = "; ".join(" - ".join(c) for c in cricche[:6])
        testo.append(f"Si individuano {len(cricche)} nuclei chiusi di almeno tre persone che si "
                     f"scelgono tutte a vicenda: {descrizione}. Sono team già coesi sul piano "
                     "relazionale.")
    else:
        testo.append("Non si rilevano nuclei chiusi di almeno tre persone legate da preferenze "
                     "reciproche: la struttura del gruppo è frammentata in coppie e legami isolati.")

    testo.append(
        "Nota metodologica: le categorie sono calcolate con il criterio di Coie e Dodge sui "
        "punteggi standardizzati di preferenza sociale e impatto sociale; su gruppi piccoli i "
        "punteggi z sono instabili, quindi le etichette vanno lette come indicazioni orientative "
        "e integrate con l'osservazione diretta.")
    return testo


def indicazioni_team(G, indici: pd.DataFrame, rec_pos: pd.DataFrame,
                     rec_neg: pd.DataFrame, cricche, nomi: dict):
    """Spunti operativi ricavati dai soli dati sociometrici."""
    voci = []
    if not rec_pos.empty:
        voci.append(("Coppie affiatate (preferenza reciproca)",
                     [f"{r.A} – {r.B}" for r in rec_pos.itertuples()]))
    if cricche:
        voci.append(("Nuclei già coesi (si scelgono tutti a vicenda)",
                     [" + ".join(c) for c in cricche]))
    if not rec_neg.empty:
        voci.append(("Accostamenti da evitare (rifiuto reciproco)",
                     [f"{r.A} – {r.B}" for r in rec_neg.itertuples()]))
    righe = []
    for ident in indici.loc[indici["Isolato"], "ID"]:
        scelti = [nomi.get(v, v) for _, v, d in G.out_edges(ident, data=True) if d["positivo"]]
        righe.append(f"{nomi.get(ident, ident)}: ha indicato "
                     f"{', '.join(scelti) if scelti else 'nessuno'}")
    if righe:
        voci.append(("Persone senza indicazioni ricevute (possibile abbinamento con chi hanno "
                     "scelto)", righe))
    return voci


def scrivi_report(percorso: Path, meta: dict, indici: pd.DataFrame, gruppo: pd.DataFrame,
                  immagini: list, rec_pos: pd.DataFrame, rec_neg: pd.DataFrame,
                  componenti, cricche, anomalie: pd.DataFrame, spunti=None):
    try:
        from docx import Document
        from docx.shared import Cm, Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except ImportError:
        print("Report .docx saltato: installa python-docx (pip install python-docx)")
        return None

    doc = Document()
    doc.add_heading("Analisi sociometrica del gruppo", level=0)
    p = doc.add_paragraph(f"File analizzato: {meta['file']} · Soggetti: {meta['n']} · "
                          f"Generato il {meta['data']}")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    doc.add_heading("1. Metodo", level=1)
    doc.add_paragraph(
        "A ciascuna persona è stato chiesto di indicare fino a tre colleghi con cui vorrebbe "
        "lavorare (SI-1, SI-2, SI-3, in ordine di preferenza) e fino a due colleghi con cui non "
        "vorrebbe lavorare (NO-1, NO-2). Le indicazioni sono state trasformate in un grafo "
        "orientato: i nodi sono le persone, gli archi le preferenze espresse. Nei conteggi "
        "ponderati la prima preferenza vale 3, la seconda 2, la terza 1; la prima indicazione "
        "negativa vale 2 e la seconda 1. Lo status sociometrico è la differenza tra preferenze "
        "positive e negative ricevute, rapportata al numero dei potenziali sceglienti, cioè delle "
        "altre persone che hanno effettivamente compilato il questionario. Le caselle lasciate "
        "vuote sono ammesse: chi compila solo in parte contribuisce con le preferenze che ha "
        "espresso, chi non compila nulla resta nel grafico come possibile destinatario di "
        "preferenze ma non entra nel denominatore degli indici.")

    doc.add_heading("2. Indici di gruppo", level=1)
    tabella = doc.add_table(rows=1, cols=2)
    tabella.style = "Light Grid Accent 1"
    tabella.rows[0].cells[0].text = "Indicatore"
    tabella.rows[0].cells[1].text = "Valore"
    for _, riga in gruppo.iterrows():
        celle = tabella.add_row().cells
        celle[0].text = str(riga["Indicatore"])
        celle[1].text = str(riga["Valore"])

    for titolo, immagine in immagini:
        doc.add_heading(titolo, level=1)
        if Path(immagine).suffix.lower() in {".png", ".jpg", ".jpeg"}:
            doc.add_picture(str(immagine), width=Cm(16))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            doc.add_paragraph(f"(grafico salvato in {Path(immagine).name})")

    doc.add_heading("5. Indici individuali", level=1)
    colonne = ["Etichetta", "Scelte ricevute", "Rifiuti ricevuti", "Scelte reciproche",
               "Status sociometrico", "Preferenza sociale (z)", "Impatto sociale (z)", "Categoria"]
    intestazioni = ["Persona", "Scelte ricevute", "Rifiuti ricevuti", "Scelte reciproche",
                    "Status", "Preferenza (z)", "Impatto (z)", "Categoria"]
    tabella = doc.add_table(rows=1, cols=len(colonne))
    tabella.style = "Light Grid Accent 1"
    for i, colonna in enumerate(intestazioni):
        cella = tabella.rows[0].cells[i]
        cella.text = colonna
        for run in cella.paragraphs[0].runs:
            run.font.size = Pt(8)
            run.bold = True
    for _, riga in indici.iterrows():
        celle = tabella.add_row().cells
        for i, colonna in enumerate(colonne):
            celle[i].text = str(riga[colonna])
            for run in celle[i].paragraphs[0].runs:
                run.font.size = Pt(8)

    doc.add_heading("6. Lettura dei risultati", level=1)
    for paragrafo in commento_automatico(indici, gruppo, componenti, cricche):
        doc.add_paragraph(paragrafo)

    doc.add_heading("7. Coppie reciproche", level=1)
    if rec_pos.empty:
        doc.add_paragraph("Nessuna scelta reciproca rilevata.")
    else:
        doc.add_paragraph("Preferenze reciproche: " +
                          "; ".join(f"{r.A} – {r.B}" for r in rec_pos.itertuples()))
    if not rec_neg.empty:
        doc.add_paragraph("Rifiuti reciproci: " +
                          "; ".join(f"{r.A} – {r.B}" for r in rec_neg.itertuples()))

    if spunti:
        doc.add_heading("8. Indicazioni per la composizione dei team", level=1)
        doc.add_paragraph(
            "Elementi ricavati esclusivamente dalle preferenze raccolte; vanno incrociati con "
            "competenze, ruoli e vincoli organizzativi prima di qualsiasi decisione.")
        for titolo, righe in spunti:
            paragrafo = doc.add_paragraph()
            paragrafo.add_run(titolo).bold = True
            for riga in righe:
                doc.add_paragraph(riga, style="List Bullet")

    doc.add_heading("9. Anomalie nei dati", level=1)
    if anomalie.empty:
        doc.add_paragraph("Nessuna anomalia rilevata nelle indicazioni raccolte.")
    else:
        for _, riga in anomalie.iterrows():
            doc.add_paragraph(f"{riga['ID']}: {riga['Tipo']} — {riga['Dettaglio']}",
                              style="List Bullet")

    doc.add_paragraph()
    nota = doc.add_paragraph(
        "Documento generato automaticamente da sociogramma.py. I commenti sono derivati dai soli "
        "dati raccolti e vanno validati da chi conduce l'analisi prima di qualsiasi utilizzo.")
    nota.runs[0].font.size = Pt(8)

    doc.save(percorso)
    return percorso


def converti_pdf(percorso_docx: Path) -> Path | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("Conversione PDF saltata: LibreOffice (soffice) non trovato.")
        return None
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir",
                    str(percorso_docx.parent), str(percorso_docx)],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    prodotto = percorso_docx.with_suffix(".pdf")
    return prodotto if prodotto.exists() else None


# ---------------------------------------------------------------------------
# File di esempio
# ---------------------------------------------------------------------------

NOMI_ESEMPIO = ["Luca", "Giulia", "Marco", "Sara", "Andrea", "Chiara", "Paolo", "Elena",
                "Davide", "Martina", "Stefano", "Alice", "Matteo", "Federica", "Simone",
                "Valentina", "Alberto", "Ilaria", "Roberto", "Silvia", "Nicola", "Laura"]
COGNOMI_ESEMPIO = ["Rossi", "Bianchi", "Ferrari", "Esposito", "Romano", "Colombo", "Ricci",
                   "Marino", "Greco", "Bruno", "Gallo", "Conti", "De Luca", "Mancini",
                   "Costa", "Giordano", "Rizzo", "Lombardi", "Moretti", "Barbieri",
                   "Fontana", "Santoro"]


def crea_esempio(percorso: Path, n: int = 22, seed: int = 11) -> Path:
    rng = np.random.default_rng(seed)
    ids = [f"S{i:02d}" for i in range(1, n + 1)]
    attrattivita = rng.gamma(2.2, 1.0, n) + 0.3
    rng.shuffle(attrattivita)
    repulsivita = (attrattivita.max() - attrattivita) + rng.gamma(1.5, 0.6, n)
    righe = []
    for i, soggetto in enumerate(ids):
        pp = attrattivita.copy(); pp[i] = 0; pp = pp / pp.sum()
        si = list(rng.choice(ids, size=3, replace=False, p=pp))
        candidati = [x for x in ids if x not in si and x != soggetto]
        pn = np.array([repulsivita[ids.index(x)] for x in candidati]); pn = pn / pn.sum()
        no = list(rng.choice(candidati, size=2, replace=False, p=pn))
        # qualche questionario incompleto o non compilato, come nella realtà
        sorte = rng.random()
        if sorte < 0.09:
            si, no = [None] * 3, [None] * 2
        elif sorte < 0.25:
            quante = int(rng.integers(1, 3))        # solo 1 o 2 preferenze su 3
            si = (si[:quante] + [None, None, None])[:3]
            no = ([no[0]] if rng.random() < 0.6 else []) + [None, None]
            no = no[:2]
        righe.append(dict(zip([COL_ID, COL_COGNOME, COL_NOME] + COL_SI + COL_NO,
                              [soggetto,
                               COGNOMI_ESEMPIO[i % len(COGNOMI_ESEMPIO)],
                               NOMI_ESEMPIO[i % len(NOMI_ESEMPIO)], *si, *no])))
    pd.DataFrame(righe).to_excel(percorso, index=False)
    return percorso


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv=None) -> int:
    p = argparse.ArgumentParser(
        description="Genera il sociogramma (diagramma di Moreno) da un Excel con "
                    "colonne ID, SI-1, SI-2, SI-3, NO-1, NO-2.")
    p.add_argument("excel", nargs="?", help="file .xlsx di input")
    p.add_argument("-o", "--output", default="output_sociogramma", help="cartella di output")
    p.add_argument("-f", "--foglio", default=0, help="nome o indice del foglio (default: il primo)")
    p.add_argument("--etichetta", metavar="COLONNA",
                   help="usa questa colonna come etichetta dei nodi al posto di NOME + COGNOME")
    p.add_argument("--etichette-dentro", action="store_true",
                   help="scrive le etichette dentro il pallino invece che sotto")
    p.add_argument("--layout", choices=["bersaglio", "spring", "kamada", "circolare"],
                   default="bersaglio", help="disposizione dei nodi (default: bersaglio)")
    p.add_argument("--insieme", action="store_true",
                   help="genera anche il grafico con scelte e rifiuti sovrapposti")
    p.add_argument("--includi-esterni", action="store_true",
                   help="tiene come nodi gli ID nominati ma assenti dalla colonna ID")
    p.add_argument("--formato", default="png", choices=["png", "pdf", "svg"],
                   help="formato dei grafici (default: png)")
    p.add_argument("--dpi", type=int, default=200)
    p.add_argument("--pdf", action="store_true",
                   help="converte anche il report in PDF (richiede LibreOffice)")
    p.add_argument("--no-html", action="store_true", help="non generare la versione interattiva")
    p.add_argument("--no-report", action="store_true", help="non generare il report .docx")
    p.add_argument("--seed", type=int, default=42, help="seed per i layout casuali")
    p.add_argument("--titolo", default="Sociogramma")
    p.add_argument("--crea-esempio", metavar="FILE.xlsx",
                   help="genera un file Excel di esempio e termina")
    args = p.parse_args(argv)

    if args.crea_esempio:
        print("File di esempio creato:", crea_esempio(Path(args.crea_esempio)))
        return 0
    if not args.excel:
        p.error("indicare il file Excel di input (oppure usare --crea-esempio)")

    percorso = Path(args.excel)
    if not percorso.exists():
        raise SystemExit(f"File non trovato: {percorso}")

    foglio = int(args.foglio) if str(args.foglio).isdigit() else args.foglio
    df, etichette, anagrafica = carica_dati(percorso, foglio, args.etichetta)
    G, anomalie = costruisci_grafo(df, includi_esterni=args.includi_esterni)
    compilazione = conteggio_compilazione(df)
    indici = calcola_indici(G, etichette, anagrafica, compilazione)
    gruppo = indici_di_gruppo(G, indici)
    componenti_ids, cricche_ids = sottogruppi(G)
    nomi = dict(zip(indici["ID"], indici["Etichetta"]))
    componenti = [[nomi.get(i, i) for i in c] for c in componenti_ids]
    cricche = [[nomi.get(i, i) for i in c] for c in cricche_ids]
    pos, anelli = calcola_layout(G, indici, args.layout, args.seed)

    out = Path(args.output)
    out.mkdir(parents=True, exist_ok=True)
    ext = args.formato
    generati = []

    dentro = True if args.etichette_dentro else None
    grafico_scelte = disegna(G, indici, pos, anelli, out / f"sociogramma_scelte.{ext}",
                             f"{args.titolo} — con chi vorrebbero lavorare", True, False,
                             args.dpi, etichette_sotto=(not dentro) if dentro else None)
    grafico_rifiuti = disegna(G, indici, pos, anelli, out / f"sociogramma_rifiuti.{ext}",
                              f"{args.titolo} — con chi non vorrebbero lavorare", False, True,
                              args.dpi, mostra_anelli=False,
                              etichette_sotto=(not dentro) if dentro else None)
    generati += [grafico_scelte, grafico_rifiuti]
    if args.insieme:
        generati.append(disegna(G, indici, pos, anelli, out / f"sociogramma_completo.{ext}",
                                f"{args.titolo} — preferenze positive e negative", True, True,
                                args.dpi, etichette_sotto=(not dentro) if dentro else None))

    if not args.no_html:
        generati.append(scrivi_html(G, indici, pos, out / "sociogramma.html",
                                    f"{args.titolo} — versione interattiva"))

    rec_pos = coppie_reciproche(G, True, nomi)
    rec_neg = coppie_reciproche(G, False, nomi)
    df_anomalie = pd.DataFrame(anomalie, columns=["ID", "Tipo", "Dettaglio"])
    foglio_gruppi = pd.DataFrame(
        [{"Tipo": "Aggregato", "Membri": " - ".join(c)} for c in componenti] +
        [{"Tipo": "Sottogruppo chiuso", "Membri": " - ".join(c)} for c in cricche],
        columns=["Tipo", "Membri"])

    report_xlsx = out / "indici.xlsx"
    with pd.ExcelWriter(report_xlsx, engine="openpyxl") as writer:
        indici.to_excel(writer, sheet_name="Indici individuali", index=False)
        gruppo.to_excel(writer, sheet_name="Indici di gruppo", index=False)
        rec_pos.to_excel(writer, sheet_name="Scelte reciproche", index=False)
        rec_neg.to_excel(writer, sheet_name="Rifiuti reciproci", index=False)
        foglio_gruppi.to_excel(writer, sheet_name="Sottogruppi", index=False)
        (indici[["ID", "Etichetta", "SI compilate", "NO compilate", "Compilazione",
                 "Potenziali sceglienti"]]
         .to_excel(writer, sheet_name="Partecipazione", index=False))
        matrice_sociometrica(G, indici).to_excel(writer, sheet_name="Matrice")
        df_anomalie.to_excel(writer, sheet_name="Anomalie", index=False)
    generati.append(report_xlsx)

    if not args.no_report:
        meta = {"file": percorso.name, "n": G.number_of_nodes(),
                "data": dt.datetime.now().strftime("%d/%m/%Y %H:%M")}
        immagini = [("3. Sociogramma delle scelte", grafico_scelte),
                    ("4. Sociogramma dei rifiuti", grafico_rifiuti)]
        spunti = indicazioni_team(G, indici, rec_pos, rec_neg, cricche, nomi)
        docx = scrivi_report(out / "report.docx", meta, indici, gruppo, immagini,
                             rec_pos, rec_neg, componenti, cricche, df_anomalie, spunti)
        if docx:
            generati.append(docx)
            if args.pdf:
                pdf = converti_pdf(docx)
                if pdf:
                    generati.append(pdf)

    n_pos = sum(1 for _, _, d in G.edges(data=True) if d["positivo"])
    n_neg = G.number_of_edges() - n_pos
    print(f"Persone: {G.number_of_nodes()} | preferenze: {n_pos} | rifiuti: {n_neg} | "
          f"anomalie: {len(anomalie)}")
    print("Categorie:", ", ".join(f"{k}={v}" for k, v in
                                  indici['Categoria'].value_counts().to_dict().items()))
    assenti = indici.loc[indici["Compilazione"] == "Assente", "Etichetta"].tolist()
    parziali = indici.loc[indici["Compilazione"] == "Parziale", "Etichetta"].tolist()
    if assenti:
        print(f"Senza preferenze espresse ({len(assenti)}):", ", ".join(assenti))
    if parziali:
        print(f"Compilazioni parziali ({len(parziali)}):", ", ".join(parziali))
    isolati = indici.loc[indici["Isolato"], "Etichetta"].tolist()
    if isolati:
        print("Senza nomine ricevute:", ", ".join(isolati))
    for f in generati:
        print("Creato:", f)
    return 0


# ---------------------------------------------------------------------------
# Interfaccia grafica (si apre facendo doppio clic sul programma)
# ---------------------------------------------------------------------------

def apri_cartella(cartella: Path) -> None:
    """Mostra la cartella dei risultati nel gestore file del sistema."""
    try:
        if sys.platform.startswith("win"):
            import os
            os.startfile(str(cartella))                      # noqa: S606 (solo Windows)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(cartella)])
        else:
            subprocess.Popen(["xdg-open", str(cartella)])
    except Exception:
        pass


AVVIO_BAT = """@echo off
cd /d "%~dp0"
title Sociogramma
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY (
  python --version >nul 2>&1 && set "PY=python"
)
if not defined PY (
  echo Python non risulta installato. Scaricalo da python.org e riprova.
  start "" https://www.python.org/downloads/windows/
  pause
  exit /b 1
)
%PY% "__PROGRAMMA__"
if errorlevel 1 pause
"""


def crea_avvio_windows(destinazione: Path) -> Path:
    """Scrive un file .bat che apre il programma con un doppio clic."""
    contenuto = AVVIO_BAT.replace("__PROGRAMMA__", str(Path(__file__).resolve()))
    percorso = destinazione / "Avvia sociogramma.bat"
    percorso.write_text(contenuto.replace("\n", "\r\n"), encoding="ascii")
    return percorso


def avvia_interfaccia() -> int:
    import queue
    import threading
    import traceback
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    coda: "queue.Queue[tuple[str, str]]" = queue.Queue()

    class _Registro:
        """Dirotta le stampe del programma nel riquadro dei messaggi."""
        def __init__(self):
            self.resto = ""

        def write(self, testo):
            self.resto += testo
            while "\n" in self.resto:
                riga, self.resto = self.resto.split("\n", 1)
                coda.put(("riga", riga))

        def flush(self):
            pass

    radice = tk.Tk()
    radice.title("Sociogramma — preferenze di collaborazione")
    radice.geometry("760x600")
    radice.minsize(680, 540)

    excel = tk.StringVar()
    cartella = tk.StringVar()
    v_report = tk.BooleanVar(value=True)
    v_html = tk.BooleanVar(value=True)
    v_pdf = tk.BooleanVar(value=False)
    v_insieme = tk.BooleanVar(value=False)

    cornice = ttk.Frame(radice, padding=14)
    cornice.pack(fill="both", expand=True)
    cornice.columnconfigure(1, weight=1)

    ttk.Label(cornice, font=("Segoe UI", 11, "bold"),
              text="Analisi delle preferenze di collaborazione").grid(
        row=0, column=0, columnspan=3, sticky="w")
    ttk.Label(cornice, foreground="#555", justify="left",
              text="Il file Excel deve avere una riga per persona e le colonne:\n"
                   "ID, COGNOME, NOME, SI-1, SI-2, SI-3, NO-1, NO-2.\n"
                   "Le caselle lasciate vuote sono ammesse.").grid(
        row=1, column=0, columnspan=3, sticky="w", pady=(4, 12))

    def scegli_file():
        percorso = filedialog.askopenfilename(
            title="Scegli il file Excel con le preferenze",
            filetypes=[("File Excel", "*.xlsx *.xlsm *.xls"), ("Tutti i file", "*.*")])
        if percorso:
            excel.set(percorso)
            if not cartella.get():
                cartella.set(str(Path(percorso).parent / "Risultati sociogramma"))

    def scegli_cartella():
        percorso = filedialog.askdirectory(title="Dove salvare i risultati")
        if percorso:
            cartella.set(percorso)

    ttk.Label(cornice, text="File Excel:").grid(row=2, column=0, sticky="w", pady=3)
    ttk.Entry(cornice, textvariable=excel).grid(row=2, column=1, sticky="ew", padx=6)
    ttk.Button(cornice, text="Sfoglia…", command=scegli_file).grid(row=2, column=2)

    ttk.Label(cornice, text="Cartella risultati:").grid(row=3, column=0, sticky="w", pady=3)
    ttk.Entry(cornice, textvariable=cartella).grid(row=3, column=1, sticky="ew", padx=6)
    ttk.Button(cornice, text="Sfoglia…", command=scegli_cartella).grid(row=3, column=2)

    opzioni = ttk.LabelFrame(cornice, text="Cosa produrre", padding=10)
    opzioni.grid(row=4, column=0, columnspan=3, sticky="ew", pady=12)
    ttk.Checkbutton(opzioni, variable=v_report,
                    text="Report Word commentato").grid(row=0, column=0, sticky="w")
    ttk.Checkbutton(opzioni, variable=v_html,
                    text="Versione interattiva da aprire nel browser").grid(row=0, column=1, sticky="w")
    ttk.Checkbutton(opzioni, variable=v_pdf,
                    text="Converti il report anche in PDF").grid(row=1, column=0, sticky="w")
    ttk.Checkbutton(opzioni, variable=v_insieme,
                    text="Aggiungi un grafico con tutto insieme").grid(row=1, column=1, sticky="w")
    ttk.Label(opzioni, foreground="#555",
              text="I due diagrammi (preferenze e rifiuti) e il foglio Excel degli indici "
                   "vengono creati sempre.").grid(row=2, column=0, columnspan=2,
                                                  sticky="w", pady=(8, 0))

    pulsanti = ttk.Frame(cornice)
    pulsanti.grid(row=5, column=0, columnspan=3, sticky="ew")
    bottone_via = ttk.Button(pulsanti, text="Genera il sociogramma")
    bottone_via.pack(side="left")
    bottone_apri = ttk.Button(pulsanti, text="Apri la cartella dei risultati", state="disabled",
                              command=lambda: apri_cartella(Path(cartella.get())))
    bottone_apri.pack(side="left", padx=8)
    ttk.Button(pulsanti, text="Crea un file Excel di esempio",
               command=lambda: crea_esempio_da_interfaccia()).pack(side="right")

    def crea_collegamento():
        cartelle = [Path.home() / "Desktop", Path.home() / "Scrivania",
                    Path(__file__).resolve().parent]
        destinazione = next((c for c in cartelle if c.is_dir()), Path.cwd())
        try:
            percorso = crea_avvio_windows(destinazione)
            messagebox.showinfo("Fatto",
                                "Ho creato il file di avvio:\n\n"
                                f"{percorso}\n\n"
                                "D'ora in poi ti basta un doppio clic su quel file per "
                                "aprire il programma.")
        except Exception as errore:
            messagebox.showerror("Non riuscito", str(errore))

    if sys.platform.startswith("win"):
        ttk.Button(pulsanti, text="Crea l'avvio rapido",
                   command=crea_collegamento).pack(side="right", padx=8)

    ttk.Label(cornice, text="Messaggi:").grid(row=6, column=0, sticky="w", pady=(12, 2))
    testo = tk.Text(cornice, height=12, wrap="word", bg="#fbfbfb")
    testo.grid(row=7, column=0, columnspan=3, sticky="nsew")
    barra = ttk.Scrollbar(cornice, command=testo.yview)
    barra.grid(row=7, column=3, sticky="ns")
    testo.configure(yscrollcommand=barra.set, state="disabled")
    cornice.rowconfigure(7, weight=1)

    def scrivi(riga: str):
        testo.configure(state="normal")
        testo.insert("end", riga + "\n")
        testo.see("end")
        testo.configure(state="disabled")

    def crea_esempio_da_interfaccia():
        percorso = filedialog.asksaveasfilename(
            title="Dove salvare il file di esempio", defaultextension=".xlsx",
            initialfile="esempio.xlsx", filetypes=[("File Excel", "*.xlsx")])
        if not percorso:
            return
        crea_esempio(Path(percorso))
        excel.set(percorso)
        if not cartella.get():
            cartella.set(str(Path(percorso).parent / "Risultati sociogramma"))
        scrivi(f"Creato il file di esempio: {percorso}")

    def esegui():
        if not excel.get():
            messagebox.showwarning("Manca il file", "Scegli prima il file Excel con le preferenze.")
            return
        if not cartella.get():
            cartella.set(str(Path(excel.get()).parent / "Risultati sociogramma"))
        bottone_via.configure(state="disabled", text="Elaborazione in corso…")
        bottone_apri.configure(state="disabled")
        testo.configure(state="normal")
        testo.delete("1.0", "end")
        testo.configure(state="disabled")

        argomenti = [excel.get(), "-o", cartella.get()]
        if not v_report.get():
            argomenti.append("--no-report")
        if not v_html.get():
            argomenti.append("--no-html")
        if v_pdf.get():
            argomenti.append("--pdf")
        if v_insieme.get():
            argomenti.append("--insieme")

        def lavoro():
            originale = sys.stdout
            sys.stdout = _Registro()
            try:
                main(argomenti)
                coda.put(("fine", ""))
            except SystemExit as errore:
                coda.put(("errore", str(errore)))
            except Exception:
                coda.put(("errore", traceback.format_exc(limit=2)))
            finally:
                sys.stdout = originale

        threading.Thread(target=lavoro, daemon=True).start()

    bottone_via.configure(command=esegui)

    def controlla():
        try:
            while True:
                tipo, contenuto = coda.get_nowait()
                if tipo == "riga":
                    scrivi(contenuto)
                elif tipo == "fine":
                    bottone_via.configure(state="normal", text="Genera il sociogramma")
                    bottone_apri.configure(state="normal")
                    scrivi("")
                    scrivi("Fatto. I file sono nella cartella dei risultati.")
                    messagebox.showinfo("Completato",
                                        "Elaborazione terminata.\n\nI file sono in:\n"
                                        f"{cartella.get()}")
                else:
                    bottone_via.configure(state="normal", text="Genera il sociogramma")
                    scrivi("ERRORE: " + contenuto)
                    messagebox.showerror("Qualcosa non ha funzionato", contenuto)
        except queue.Empty:
            pass
        radice.after(120, controlla)

    radice.after(120, controlla)
    radice.mainloop()
    return 0


if __name__ == "__main__":
    if len(sys.argv) == 1:
        sys.exit(avvia_interfaccia())
    sys.exit(main())
