"""
╔══════════════════════════════════════════════════════════════════╗
║  SISTEMA DE NOTIFICAÇÕES — Etapa 4.6                            ║
║  Notificações automáticas do funcionário                        ║
║  · Barra de notificações do Windows (toast nativo)              ║
║  · Central de notificações dentro do app                        ║
║  · Sistema inteligente anti-spam com agrupamento                ║
╚══════════════════════════════════════════════════════════════════╝
"""

import sqlite3
import threading
import time
import subprocess
import sys
import os
import json
from datetime import datetime, timedelta
from typing import Optional, Callable

# ── Tkinter (opcional para uso standalone) ───────────────────────────────────
try:
    import tkinter as tk
    from tkinter import ttk
    _TK_OK = True
except ImportError:
    _TK_OK = False

# ── Cores e fontes (espelha ponto.py) ────────────────────────────────────────
C = {
    "bg_dark":    "#0D0F14",  "bg_panel":   "#13161E",
    "bg_card":    "#1A1E2A",  "bg_input":   "#0F1219",
    "sidebar":    "#0A0C12",  "accent":     "#00C2FF",
    "accent_dim": "#005599",  "success":    "#00E5A0",
    "warning":    "#FFB830",  "danger":     "#FF4D6A",
    "purple":     "#AA88FF",  "text_hi":    "#EAEDF5",
    "text_mid":   "#8A90A8",  "text_lo":    "#4A5068",
    "border":     "#1E2235",  "hover":      "#1F2438",
    "white":      "#FFFFFF",
}
FONT_SMALL  = ("Segoe UI",  9)
FONT_BODY   = ("Segoe UI", 11)
FONT_HEADER = ("Segoe UI", 13, "bold")


# ══════════════════════════════════════════════════════════════════════════════
# TIPOS DE NOTIFICAÇÃO
# ══════════════════════════════════════════════════════════════════════════════
class TipoNotif:
    HORAS_EXTRAS     = "horas_extras"      # 🟢
    SALDO_ATUALIZADO = "saldo_atualizado"  # 💰
    HORAS_NEGATIVAS  = "horas_negativas"   # ⚠
    FECHAMENTO       = "fechamento"        # 📊
    PAGAMENTO        = "pagamento"         # 💵

    ICONES = {
        "horas_extras":     "🟢",
        "saldo_atualizado": "💰",
        "horas_negativas":  "⚠",
        "fechamento":       "📊",
        "pagamento":        "💵",
    }
    CORES = {
        "horas_extras":     "#00E5A0",   # success
        "saldo_atualizado": "#FFB830",   # warning (ouro)
        "horas_negativas":  "#FF4D6A",   # danger
        "fechamento":       "#00C2FF",   # accent
        "pagamento":        "#AA88FF",   # purple
    }
    # Cooldown mínimo entre notificações do mesmo tipo/func (em minutos)
    COOLDOWN = {
        "horas_extras":     60,
        "saldo_atualizado": 120,
        "horas_negativas":  90,
        "fechamento":       0,    # sempre enviar
        "pagamento":        0,    # sempre enviar
    }


# ══════════════════════════════════════════════════════════════════════════════
# BANCO DE DADOS DE NOTIFICAÇÕES
# ══════════════════════════════════════════════════════════════════════════════
class NotifDB:
    """Gerencia a tabela de notificações no mesmo arquivo ponto.db."""

    def __init__(self, conn: sqlite3.Connection):
        self.conn = conn
        self._migrate()

    def _migrate(self):
        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS notificacoes (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            func_id     INTEGER,
            tipo        TEXT    NOT NULL,
            titulo      TEXT    NOT NULL,
            mensagem    TEXT    NOT NULL,
            navegacao   TEXT,
            lida        INTEGER NOT NULL DEFAULT 0,
            enviada_win INTEGER NOT NULL DEFAULT 0,
            criado_em   TEXT    NOT NULL
        );

        CREATE TABLE IF NOT EXISTS notif_cooldown (
            func_id   INTEGER NOT NULL,
            tipo      TEXT    NOT NULL,
            ultima_em TEXT    NOT NULL,
            PRIMARY KEY (func_id, tipo)
        );
        """)
        self.conn.commit()

    # ── CRUD ─────────────────────────────────────────────────────────────────

    def inserir(self, func_id: Optional[int], tipo: str, titulo: str,
                mensagem: str, navegacao: str = "") -> int:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur = self.conn.execute(
            "INSERT INTO notificacoes(func_id,tipo,titulo,mensagem,navegacao,criado_em)"
            " VALUES(?,?,?,?,?,?)",
            (func_id, tipo, titulo, mensagem, navegacao, now))
        self.conn.commit()
        return cur.lastrowid

    def listar(self, func_id: Optional[int] = None, apenas_nao_lidas: bool = False,
               limite: int = 80) -> list:
        q = "SELECT * FROM notificacoes WHERE 1=1"
        params = []
        if func_id is not None:
            q += " AND (func_id=? OR func_id IS NULL)"
            params.append(func_id)
        if apenas_nao_lidas:
            q += " AND lida=0"
        q += " ORDER BY criado_em DESC LIMIT ?"
        params.append(limite)
        return self.conn.execute(q, params).fetchall()

    def marcar_lida(self, notif_id: int):
        self.conn.execute("UPDATE notificacoes SET lida=1 WHERE id=?", (notif_id,))
        self.conn.commit()

    def marcar_todas_lidas(self, func_id: Optional[int] = None):
        if func_id is not None:
            self.conn.execute(
                "UPDATE notificacoes SET lida=1 WHERE func_id=? OR func_id IS NULL",
                (func_id,))
        else:
            self.conn.execute("UPDATE notificacoes SET lida=1")
        self.conn.commit()

    def contar_nao_lidas(self, func_id: Optional[int] = None) -> int:
        q = "SELECT COUNT(*) FROM notificacoes WHERE lida=0"
        params = []
        if func_id is not None:
            q += " AND (func_id=? OR func_id IS NULL)"
            params.append(func_id)
        return self.conn.execute(q, params).fetchone()[0]

    def deletar_antigas(self, dias: int = 30):
        limite = (datetime.now() - timedelta(days=dias)).strftime("%Y-%m-%d %H:%M:%S")
        self.conn.execute("DELETE FROM notificacoes WHERE criado_em < ? AND lida=1", (limite,))
        self.conn.commit()

    # ── Cooldown ─────────────────────────────────────────────────────────────

    def pode_enviar(self, func_id: int, tipo: str) -> bool:
        cooldown_min = TipoNotif.COOLDOWN.get(tipo, 60)
        if cooldown_min == 0:
            return True
        row = self.conn.execute(
            "SELECT ultima_em FROM notif_cooldown WHERE func_id=? AND tipo=?",
            (func_id, tipo)).fetchone()
        if not row:
            return True
        try:
            ultima = datetime.strptime(row["ultima_em"], "%Y-%m-%d %H:%M:%S")
            return (datetime.now() - ultima).total_seconds() > cooldown_min * 60
        except Exception:
            return True

    def registrar_envio(self, func_id: int, tipo: str):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.conn.execute(
            "INSERT OR REPLACE INTO notif_cooldown(func_id,tipo,ultima_em) VALUES(?,?,?)",
            (func_id, tipo, now))
        self.conn.commit()

    def marcar_enviada_win(self, notif_id: int):
        self.conn.execute("UPDATE notificacoes SET enviada_win=1 WHERE id=?", (notif_id,))
        self.conn.commit()


# ══════════════════════════════════════════════════════════════════════════════
# TOAST WINDOWS (PowerShell / balão tkinter fallback)
# ══════════════════════════════════════════════════════════════════════════════
def _toast_windows(titulo: str, mensagem: str, tipo: str = ""):
    """
    Envia notificação toast nativa do Windows via PowerShell.
    Funciona no Windows 10/11. Silencia erros — nunca trava o app.
    """
    try:
        icone_map = {
            "horas_extras":     "Info",
            "saldo_atualizado": "Info",
            "horas_negativas":  "Warning",
            "fechamento":       "Info",
            "pagamento":        "Info",
        }
        icone = icone_map.get(tipo, "Info")

        # Script PowerShell para toast nativo Windows 10/11
        ps_script = f"""
Add-Type -AssemblyName System.Windows.Forms
$notify = New-Object System.Windows.Forms.NotifyIcon
$notify.Icon = [System.Drawing.SystemIcons]::{icone}
$notify.Visible = $true
$notify.ShowBalloonTip(6000, '{titulo.replace("'", "")}', '{mensagem.replace("'", "")}', [System.Windows.Forms.ToolTipIcon]::{icone})
Start-Sleep -Milliseconds 6500
$notify.Dispose()
""".strip()

        subprocess.Popen(
            ["powershell", "-WindowStyle", "Hidden", "-NonInteractive",
             "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=0x08000000  # CREATE_NO_WINDOW
        )
    except Exception:
        pass  # Silencia — ambiente não-Windows ou sem PowerShell


# ══════════════════════════════════════════════════════════════════════════════
# GERENCIADOR CENTRAL DE NOTIFICAÇÕES
# ══════════════════════════════════════════════════════════════════════════════
class GerenciadorNotif:
    """
    Ponto central para emitir notificações.
    · Verifica cooldown antes de enviar
    · Persiste no banco
    · Dispara toast Windows em thread separada
    · Chama callbacks de UI registrados
    """

    def __init__(self, conn: sqlite3.Connection):
        self.db = NotifDB(conn)
        self._callbacks: list[Callable] = []   # chamados quando nova notif chega
        self._lock = threading.Lock()

    def registrar_callback(self, fn: Callable):
        """Registra função chamada após cada nova notificação (atualiza badge, etc.)."""
        self._callbacks.append(fn)

    def _notificar_callbacks(self):
        for fn in self._callbacks:
            try:
                fn()
            except Exception:
                pass

    # ── API pública ───────────────────────────────────────────────────────────

    def emitir(self, tipo: str, titulo: str, mensagem: str,
               func_id: Optional[int] = None, navegacao: str = "",
               forcar: bool = False) -> Optional[int]:
        """
        Emite uma notificação inteligente.
        - forcar=True ignora cooldown (ex: pagamentos, fechamentos)
        - Retorna o ID da notificação inserida, ou None se bloqueada por cooldown.
        """
        with self._lock:
            fid_cooldown = func_id if func_id is not None else 0

            if not forcar and not self.db.pode_enviar(fid_cooldown, tipo):
                return None   # bloqueada por cooldown

            # Persiste no banco
            notif_id = self.db.inserir(func_id, tipo, titulo, mensagem, navegacao)

            # Toast Windows em thread para não travar UI
            threading.Thread(
                target=self._enviar_toast,
                args=(notif_id, titulo, mensagem, tipo),
                daemon=True
            ).start()

            # Registra cooldown
            if not forcar:
                self.db.registrar_envio(fid_cooldown, tipo)

        # Avisa callbacks de UI (fora do lock)
        self._notificar_callbacks()
        return notif_id

    def _enviar_toast(self, notif_id: int, titulo: str, mensagem: str, tipo: str):
        _toast_windows(titulo, mensagem, tipo)
        self.db.marcar_enviada_win(notif_id)

    # ── Métodos de conveniência ───────────────────────────────────────────────

    def horas_extras(self, func_nome: str, func_id: int,
                     horas_str: str, data_str: str = ""):
        """Ex: horas_str = '+2h14'"""
        titulo = "Horas Extras Calculadas"
        msg = f"Você fez {horas_str} de horas extras"
        if data_str:
            msg += f" em {data_str}"
        msg += "."
        nav = f"calculo:{func_id}:{data_str}" if data_str else f"calculo:{func_id}"
        return self.emitir(TipoNotif.HORAS_EXTRAS, titulo, msg, func_id, nav)

    def saldo_atualizado(self, func_nome: str, func_id: int, valor_brl: str):
        """Ex: valor_brl = 'R$ 84,08'"""
        titulo = "Saldo de Horas Atualizado"
        msg = f"Seu saldo atual de horas extras é {valor_brl}."
        return self.emitir(TipoNotif.SALDO_ATUALIZADO, titulo, msg, func_id,
                           f"financeiro:{func_id}")

    def horas_negativas(self, func_nome: str, func_id: int, tempo_str: str):
        """Ex: tempo_str = '20 minutos'"""
        titulo = "Atenção: Horas Negativas"
        msg = f"Você possui {tempo_str} negativos."
        return self.emitir(TipoNotif.HORAS_NEGATIVAS, titulo, msg, func_id,
                           f"calculo:{func_id}", forcar=True)

    def fechamento_semanal(self, semana_str: str = ""):
        """Semana fechada — notificação global."""
        titulo = "Fechamento Semanal"
        msg = "Fechamento semanal concluído."
        if semana_str:
            msg = f"Fechamento da semana {semana_str} concluído."
        return self.emitir(TipoNotif.FECHAMENTO, titulo, msg, None,
                           "banco_horas", forcar=True)

    def pagamento_lancado(self, func_nome: str, func_id: int, valor_brl: str):
        titulo = "Pagamento Lançado"
        msg = f"Pagamento de {valor_brl} registrado para {func_nome}."
        return self.emitir(TipoNotif.PAGAMENTO, titulo, msg, func_id,
                           f"financeiro:{func_id}", forcar=True)


# ══════════════════════════════════════════════════════════════════════════════
# WIDGET — BADGE (bolinha vermelha com contagem)
# ══════════════════════════════════════════════════════════════════════════════
class BadgeNotif(tk.Frame):
    """
    Pequeno badge animado com contagem de não-lidas.
    Colado ao lado do botão de sino na sidebar/header.
    """

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=C["sidebar"], **kw)
        self._count = 0
        self._lbl = tk.Label(self, text="", bg=C["danger"], fg=C["white"],
                             font=("Segoe UI", 7, "bold"),
                             width=2, padx=3, pady=1)
        # Não exibido até ter notificações
        self._visible = False

    def atualizar(self, count: int):
        self._count = count
        if count > 0:
            self._lbl.config(text=str(count) if count < 100 else "99+")
            if not self._visible:
                self._lbl.pack()
                self._visible = True
            self._animar()
        else:
            self._lbl.pack_forget()
            self._visible = False

    def _animar(self):
        """Pulsa suavemente o badge."""
        colors = [C["danger"], "#FF7A8A", C["danger"]]
        def _step(i=0):
            if i < len(colors) and self._visible:
                self._lbl.config(bg=colors[i])
                self.after(180, lambda: _step(i + 1))
        _step()


# ══════════════════════════════════════════════════════════════════════════════
# WIDGET — CENTRAL DE NOTIFICAÇÕES (janela dentro do app)
# ══════════════════════════════════════════════════════════════════════════════
class CentralNotificacoes(tk.Toplevel):
    """
    Painel flutuante com lista completa de notificações.
    Abre ao clicar no ícone de sino.
    """

    def __init__(self, parent, gerenciador: GerenciadorNotif,
                 func_id: Optional[int] = None,
                 on_navegar: Optional[Callable] = None):
        super().__init__(parent)
        self.gerenciador = gerenciador
        self.func_id     = func_id
        self.on_navegar  = on_navegar

        self.title("Central de Notificações")
        self.configure(bg=C["bg_dark"])
        self.resizable(True, True)
        self.minsize(480, 520)

        # Centraliza sobre o parent
        px = parent.winfo_rootx(); py = parent.winfo_rooty()
        pw = parent.winfo_width(); ph = parent.winfo_height()
        self.geometry(f"520x620+{px + pw//2 - 260}+{py + ph//2 - 310}")

        self.transient(parent)
        self._build()
        self._carregar()

    # ── Construção UI ─────────────────────────────────────────────────────────

    def _build(self):
        # Header
        hdr = tk.Frame(self, bg=C["bg_dark"], padx=20, pady=16)
        hdr.pack(fill="x")

        tk.Label(hdr, text="🔔  Central de Notificações",
                 bg=C["bg_dark"], fg=C["text_hi"],
                 font=("Segoe UI", 14, "bold")).pack(side="left")

        btn_lidas = tk.Label(hdr, text="Marcar todas como lidas",
                             bg=C["bg_dark"], fg=C["accent"],
                             font=("Segoe UI", 9), cursor="hand2")
        btn_lidas.pack(side="right", padx=4)
        btn_lidas.bind("<Button-1>", self._marcar_todas_lidas)

        # Separador
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        # Filtro rápido
        filtro_frame = tk.Frame(self, bg=C["bg_panel"], padx=16, pady=10)
        filtro_frame.pack(fill="x")

        self._filtro = tk.StringVar(value="todas")
        for val, txt in [("todas", "Todas"), ("nao_lidas", "Não lidas"),
                         ("horas_extras", "🟢 Extras"),
                         ("horas_negativas", "⚠ Negativas"),
                         ("pagamento", "💵 Pagamentos")]:
            rb = tk.Radiobutton(filtro_frame, text=txt, variable=self._filtro,
                                value=val, bg=C["bg_panel"], fg=C["text_mid"],
                                activebackground=C["bg_panel"],
                                selectcolor=C["bg_card"],
                                font=("Segoe UI", 9),
                                command=self._carregar)
            rb.pack(side="left", padx=6)

        # Separador
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        # Lista scrollável
        container = tk.Frame(self, bg=C["bg_panel"])
        container.pack(fill="both", expand=True)

        self._canvas = tk.Canvas(container, bg=C["bg_panel"],
                                 highlightthickness=0, bd=0)
        sb = ttk.Scrollbar(container, orient="vertical",
                           command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=sb.set)

        sb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._lista = tk.Frame(self._canvas, bg=C["bg_panel"])
        self._window = self._canvas.create_window((0, 0), window=self._lista,
                                                  anchor="nw")

        self._lista.bind("<Configure>",
                         lambda e: self._canvas.configure(
                             scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
                          lambda e: self._canvas.itemconfig(
                              self._window, width=e.width))
        # Scroll com roda do mouse
        self._canvas.bind_all("<MouseWheel>",
                              lambda e: self._canvas.yview_scroll(
                                  -1 * (e.delta // 120), "units"))

        # Rodapé
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")
        rod = tk.Frame(self, bg=C["bg_dark"], padx=16, pady=10)
        rod.pack(fill="x")
        self._status_lbl = tk.Label(rod, text="", bg=C["bg_dark"],
                                    fg=C["text_lo"], font=FONT_SMALL)
        self._status_lbl.pack(side="left")

    # ── Carregamento ──────────────────────────────────────────────────────────

    def _carregar(self):
        for w in self._lista.winfo_children():
            w.destroy()

        filtro = self._filtro.get()
        apenas_nao_lidas = (filtro == "nao_lidas")
        notifs = self.gerenciador.db.listar(
            func_id=self.func_id, apenas_nao_lidas=apenas_nao_lidas)

        # Filtra por tipo se necessário
        if filtro not in ("todas", "nao_lidas"):
            notifs = [n for n in notifs if n["tipo"] == filtro]

        if not notifs:
            vazio = tk.Frame(self._lista, bg=C["bg_panel"], pady=60)
            vazio.pack(fill="x")
            tk.Label(vazio, text="🔕", bg=C["bg_panel"], fg=C["text_lo"],
                     font=("Segoe UI", 32)).pack()
            tk.Label(vazio, text="Nenhuma notificação",
                     bg=C["bg_panel"], fg=C["text_lo"],
                     font=("Segoe UI", 12)).pack(pady=8)
            self._status_lbl.config(text="0 notificações")
            return

        nao_lidas = sum(1 for n in notifs if not n["lida"])
        self._status_lbl.config(
            text=f"{len(notifs)} notificações  ·  {nao_lidas} não lidas")

        # Agrupa por dia
        grupos = {}
        for n in notifs:
            try:
                dt = datetime.strptime(n["criado_em"], "%Y-%m-%d %H:%M:%S")
                dia = dt.strftime("%d/%m/%Y")
            except Exception:
                dia = "—"
            grupos.setdefault(dia, []).append(n)

        hoje = datetime.now().strftime("%d/%m/%Y")
        ontem = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
        label_dia = {hoje: "Hoje", ontem: "Ontem"}

        for dia, items in grupos.items():
            # Cabeçalho do dia
            dh = tk.Frame(self._lista, bg=C["bg_panel"], padx=16, pady=8)
            dh.pack(fill="x")
            tk.Label(dh, text=label_dia.get(dia, dia),
                     bg=C["bg_panel"], fg=C["text_lo"],
                     font=("Segoe UI", 9, "bold")).pack(side="left")
            tk.Frame(dh, bg=C["border"], height=1).pack(
                side="left", fill="x", expand=True, padx=8)

            for n in items:
                self._card_notif(n)

    def _card_notif(self, n):
        """Renderiza um card de notificação com animação de entrada."""
        tipo  = n["tipo"]
        lida  = bool(n["lida"])
        cor   = TipoNotif.CORES.get(tipo, C["accent"])
        icone = TipoNotif.ICONES.get(tipo, "🔔")
        nav   = n["navegacao"] or ""

        bg_card = C["bg_card"] if not lida else C["bg_panel"]

        outer = tk.Frame(self._lista, bg=C["border"], pady=1)
        outer.pack(fill="x", padx=0, pady=1)

        card = tk.Frame(outer, bg=bg_card, padx=16, pady=12, cursor="hand2")
        card.pack(fill="x")

        # Barra colorida lateral
        barra = tk.Frame(card, bg=cor if not lida else C["border"], width=3)
        barra.pack(side="left", fill="y", padx=(0, 12))

        # Conteúdo
        corpo = tk.Frame(card, bg=bg_card)
        corpo.pack(side="left", fill="both", expand=True)

        # Linha superior: ícone + título + hora
        topo = tk.Frame(corpo, bg=bg_card)
        topo.pack(fill="x")

        tk.Label(topo, text=icone, bg=bg_card, fg=cor,
                 font=("Segoe UI", 13)).pack(side="left", padx=(0, 6))

        tk.Label(topo, text=n["titulo"], bg=bg_card,
                 fg=C["text_hi"] if not lida else C["text_mid"],
                 font=("Segoe UI", 11, "bold" if not lida else "normal"),
                 anchor="w").pack(side="left", fill="x", expand=True)

        try:
            hora = datetime.strptime(
                n["criado_em"], "%Y-%m-%d %H:%M:%S").strftime("%H:%M")
        except Exception:
            hora = ""
        tk.Label(topo, text=hora, bg=bg_card, fg=C["text_lo"],
                 font=FONT_SMALL).pack(side="right")

        # Mensagem
        tk.Label(corpo, text=n["mensagem"], bg=bg_card,
                 fg=C["text_mid"], font=FONT_SMALL,
                 anchor="w", justify="left",
                 wraplength=360).pack(fill="x", pady=(4, 0))

        # Link de navegação
        if nav:
            nav_lbl = tk.Label(corpo, text="Ver detalhes →",
                               bg=bg_card, fg=cor,
                               font=("Segoe UI", 9), cursor="hand2")
            nav_lbl.pack(anchor="w", pady=(4, 0))
            nav_lbl.bind("<Button-1>",
                         lambda e, nid=n["id"], dest=nav: self._navegar(nid, dest))

        # Badge "novo"
        if not lida:
            tk.Label(card, text="●", bg=bg_card, fg=cor,
                     font=("Segoe UI", 8)).pack(side="right", padx=4)

        # Hover
        def _enter(e, f=card, b=bg_card):
            f.config(bg=C["hover"])
            for child in f.winfo_children():
                _recolor(child, C["hover"])

        def _leave(e, f=card, b=bg_card):
            f.config(bg=b)
            for child in f.winfo_children():
                _recolor(child, b)

        def _recolor(widget, bg):
            try:
                widget.config(bg=bg)
                for c in widget.winfo_children():
                    _recolor(c, bg)
            except Exception:
                pass

        card.bind("<Enter>", _enter)
        card.bind("<Leave>", _leave)
        card.bind("<Button-1>",
                  lambda e, nid=n["id"], dest=nav: self._navegar(nid, dest))

    # ── Ações ─────────────────────────────────────────────────────────────────

    def _navegar(self, notif_id: int, destino: str):
        self.gerenciador.db.marcar_lida(notif_id)
        self._carregar()
        if self.on_navegar and destino:
            self.on_navegar(destino)
        self.withdraw()

    def _marcar_todas_lidas(self, e=None):
        self.gerenciador.db.marcar_todas_lidas(self.func_id)
        self._carregar()
        self.gerenciador._notificar_callbacks()


# ══════════════════════════════════════════════════════════════════════════════
# WIDGET — BOTÃO DE SINO COM BADGE (para sidebar do app principal)
# ══════════════════════════════════════════════════════════════════════════════
class BotaoSino(tk.Frame):
    """
    Botão 🔔 com badge de contagem de não-lidas.
    Conecta ao GerenciadorNotif e abre CentralNotificacoes ao clicar.
    """

    def __init__(self, parent, gerenciador: GerenciadorNotif,
                 func_id: Optional[int] = None,
                 on_navegar: Optional[Callable] = None,
                 bg_color: str = C["sidebar"], **kw):
        super().__init__(parent, bg=bg_color, cursor="hand2", **kw)
        self.gerenciador = gerenciador
        self.func_id     = func_id
        self.on_navegar  = on_navegar
        self._central    = None
        self._bg         = bg_color
        self._piscar_id  = None

        self._build()
        gerenciador.registrar_callback(self._atualizar_badge)
        self._atualizar_badge()

    def _build(self):
        # Frame interno para ícone + badge
        inner = tk.Frame(self, bg=self._bg)
        inner.pack(padx=8, pady=8)

        self._sino = tk.Label(inner, text="🔔", bg=self._bg, fg=C["text_mid"],
                              font=("Segoe UI", 15))
        self._sino.pack(side="left")

        self._badge = tk.Label(inner, text="", bg=C["danger"], fg=C["white"],
                               font=("Segoe UI", 7, "bold"), padx=4, pady=1)
        # Não exibido ainda

        # Eventos
        for w in (self, inner, self._sino):
            w.bind("<Button-1>", self._abrir)
            w.bind("<Enter>",    lambda e: self._sino.config(fg=C["accent"]))
            w.bind("<Leave>",    lambda e: self._sino.config(fg=C["text_mid"]))

    def _atualizar_badge(self):
        count = self.gerenciador.db.contar_nao_lidas(self.func_id)
        if count > 0:
            self._badge.config(text=str(count) if count < 100 else "99+")
            self._badge.pack(side="left")
            self._iniciar_piscar()
        else:
            self._badge.pack_forget()
            self._parar_piscar()

    def _iniciar_piscar(self):
        if self._piscar_id:
            return
        self._piscar(True)

    def _parar_piscar(self):
        if self._piscar_id:
            self.after_cancel(self._piscar_id)
            self._piscar_id = None
        self._sino.config(fg=C["text_mid"])

    def _piscar(self, estado: bool):
        self._sino.config(fg=C["accent"] if estado else C["text_mid"])
        self._piscar_id = self.after(1200, lambda: self._piscar(not estado))

    def _abrir(self, e=None):
        if self._central and self._central.winfo_exists():
            self._central.lift()
            self._central.focus_force()
            return
        # Encontra janela raiz
        root = self.winfo_toplevel()
        self._central = CentralNotificacoes(
            root, self.gerenciador,
            func_id=self.func_id,
            on_navegar=self.on_navegar)
        self._central.protocol("WM_DELETE_WINDOW", self._central.withdraw)


# ══════════════════════════════════════════════════════════════════════════════
# POPUP TOAST INTERNO (aparece no canto do app, independente do Windows)
# ══════════════════════════════════════════════════════════════════════════════
class ToastInterno:
    """
    Toast que aparece no canto inferior direito da janela principal do app.
    Usa after() do tkinter — thread-safe e sem janela extra do SO.
    """

    _fila: list = []
    _mostrando: bool = False

    @classmethod
    def mostrar(cls, root: tk.Tk, titulo: str, mensagem: str,
                tipo: str = "", duracao_ms: int = 5000,
                on_click: Optional[Callable] = None):
        cls._fila.append((root, titulo, mensagem, tipo, duracao_ms, on_click))
        if not cls._mostrando:
            cls._proximo(root)

    @classmethod
    def _proximo(cls, root: tk.Tk):
        if not cls._fila:
            cls._mostrando = False
            return
        cls._mostrando = True
        args = cls._fila.pop(0)
        cls._exibir(*args)

    @classmethod
    def _exibir(cls, root, titulo, mensagem, tipo, duracao_ms, on_click):
        cor  = TipoNotif.CORES.get(tipo, C["accent"])
        icon = TipoNotif.ICONES.get(tipo, "🔔")

        win = tk.Toplevel(root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.configure(bg=C["bg_card"])

        # Sombra / borda
        borda = tk.Frame(win, bg=cor, padx=1, pady=1)
        borda.pack(fill="both", expand=True)
        inner = tk.Frame(borda, bg=C["bg_card"], padx=16, pady=12)
        inner.pack(fill="both", expand=True)

        # Barra colorida no topo
        tk.Frame(inner, bg=cor, height=2).pack(fill="x")

        # Ícone + título
        topo = tk.Frame(inner, bg=C["bg_card"], pady=6)
        topo.pack(fill="x")
        tk.Label(topo, text=icon, bg=C["bg_card"], fg=cor,
                 font=("Segoe UI", 14)).pack(side="left", padx=(0, 8))
        tk.Label(topo, text=titulo, bg=C["bg_card"], fg=C["text_hi"],
                 font=("Segoe UI", 10, "bold"), anchor="w").pack(
                     side="left", fill="x", expand=True)

        # Fechar
        def _fechar(e=None):
            try:
                win.destroy()
            except Exception:
                pass
            root.after(400, lambda: cls._proximo(root))

        tk.Label(topo, text="✕", bg=C["bg_card"], fg=C["text_lo"],
                 font=("Segoe UI", 9), cursor="hand2").pack(
                     side="right").bind("<Button-1>", _fechar)

        # Mensagem
        tk.Label(inner, text=mensagem, bg=C["bg_card"], fg=C["text_mid"],
                 font=FONT_SMALL, anchor="w", justify="left",
                 wraplength=280).pack(fill="x", pady=(0, 6))

        # Barra de progresso animada
        prog_outer = tk.Frame(inner, bg=C["border"], height=2)
        prog_outer.pack(fill="x")
        prog_bar = tk.Frame(prog_outer, bg=cor, height=2)
        prog_bar.place(relwidth=1.0, relheight=1.0)

        # Eventos de clique
        if on_click:
            for w in (inner, topo):
                w.bind("<Button-1>", lambda e: (on_click(), _fechar()))
                w.config(cursor="hand2")

        # Posicionar no canto inferior direito do root
        root.update_idletasks()
        rx = root.winfo_x(); ry = root.winfo_y()
        rw = root.winfo_width(); rh = root.winfo_height()
        win.update_idletasks()
        ww = win.winfo_reqwidth() or 320
        wh = win.winfo_reqheight() or 100

        # Garante largura mínima
        win.geometry(f"320x{max(80, wh)}+{rx + rw - 340}+{ry + rh - max(80,wh) - 20}")

        # Animação: slide de entrada
        target_y = ry + rh - max(80, wh) - 20
        start_y  = target_y + 60

        def _slide(y):
            if y > target_y:
                try:
                    win.geometry(
                        f"320x{max(80,wh)}+{rx + rw - 340}+{y}")
                    root.after(16, lambda: _slide(y - 5))
                except Exception:
                    pass

        _slide(start_y)

        # Anima barra de progresso
        steps = duracao_ms // 50
        def _prog(step):
            if step <= 0:
                _fechar()
                return
            try:
                rel = step / steps
                prog_bar.place(relwidth=rel, relheight=1.0)
                root.after(50, lambda: _prog(step - 1))
            except Exception:
                pass
        _prog(steps)
