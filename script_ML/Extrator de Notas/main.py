import threading
import webbrowser
import tkinter as tk
from tkinter import messagebox, simpledialog
from datetime import datetime, timedelta

from config import get_appdata_dir, COOKIE_PATH
import auth_ml
import core

# ── Paleta UI ─────────────────────────────────────────────────────────────────
BG, SURFACE, BORDER = "#0f1117", "#1a1d27", "#2a2f42"
ACCENT, GREEN, RED = "#f5a623", "#39d98a", "#ff5c5c"
TEXT, MUTED = "#e8eaf0", "#6b7280"
MONO, SMALL = ("Courier New", 10), ("Segoe UI", 9)

class ConfigWindow(tk.Toplevel):
    def __init__(self, parent, cfg: dict, on_save):
        super().__init__(parent)
        self.title("Configurações"); self.configure(bg=BG); self.resizable(False, False)
        self.geometry(f"480x360+{parent.winfo_x()+50}+{parent.winfo_y()+50}"); self.grab_set()
        self.on_save = on_save; self.vars = {}
        
        p = tk.Frame(self, bg=BG, padx=24, pady=20); p.pack(fill="both", expand=True)
        tk.Label(p, text="Configurações", font=("Segoe UI", 12, "bold"), fg=ACCENT, bg=BG).pack(anchor="w", pady=(0,14))
        
        for key, label, secret in [("redirect_uri", "Redirect URI", False), ("client_id", "Client ID", False), ("client_secret", "Client Secret", True)]:
            tk.Label(p, text=label, font=("Courier New", 8, "bold"), fg=MUTED, bg=BG, anchor="w").pack(fill="x")
            v = tk.StringVar(value=cfg.get(key, ""))
            self.vars[key] = v
            tk.Entry(p, textvariable=v, show="●" if secret else "", bg="#0a0c12", fg=TEXT, insertbackground=ACCENT, relief="flat", font=MONO).pack(fill="x", ipady=5, ipadx=6, pady=(2,10))

        btn = tk.Frame(p, bg=BG); btn.pack(fill="x", pady=(4,0))
        tk.Button(btn, text="Fechar", font=SMALL, bg=SURFACE, fg=MUTED, relief="flat", command=self.destroy).pack(side="right", padx=(8,0))
        tk.Button(btn, text="Salvar", font=("Segoe UI", 9, "bold"), bg=ACCENT, fg=BG, relief="flat", command=self._save).pack(side="right")

    def _save(self):
        self.on_save({k: v.get().strip() for k, v in self.vars.items()})
        self.destroy()

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ML ERP Extractor Pro")
        self.configure(bg=BG)
        self.geometry("640x550")
        self.cfg = auth_ml.load_config()
        self.auth = auth_ml.load_auth()
        self._build()

    def _build(self):
        # ── Topo ──
        top = tk.Frame(self, bg=SURFACE, padx=16, pady=10)
        top.pack(fill="x")
        tk.Button(top, text="⚙ Configurações", font=SMALL, bg=SURFACE, fg=MUTED, relief="flat", command=lambda: ConfigWindow(self, self.cfg, self._on_config_saved)).pack(side="left")
        tk.Button(top, text="🔑 Autenticar", font=SMALL, bg=SURFACE, fg=MUTED, relief="flat", command=self._start_oauth).pack(side="left", padx=8)
        tk.Button(top, text="✕ Desconectar", font=SMALL, bg=SURFACE, fg=RED, relief="flat", command=self._revoke).pack(side="left")
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        # ── Centro ──
        center = tk.Frame(self, bg=BG, padx=32, pady=20)
        center.pack(fill="both", expand=True)

        # Datas
        dates = tk.Frame(center, bg=BG); dates.pack(fill="x", pady=(0,15))
        tk.Label(dates, text="PERÍODO DE EXTRAÇÃO", font=("Courier New", 8, "bold"), fg=MUTED, bg=BG).pack(anchor="w", pady=(0, 6))
        row = tk.Frame(dates, bg=BG); row.pack(fill="x")
        
        self.v_inicio = tk.StringVar(value=(datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d"))
        self.v_fim = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        
        tk.Label(row, text="De", font=SMALL, fg=MUTED, bg=BG).pack(side="left")
        tk.Entry(row, textvariable=self.v_inicio, width=12, bg="#0a0c12", fg=TEXT, font=MONO).pack(side="left", padx=6, ipady=3)
        tk.Label(row, text="até", font=SMALL, fg=MUTED, bg=BG).pack(side="left")
        tk.Entry(row, textvariable=self.v_fim, width=12, bg="#0a0c12", fg=TEXT, font=MONO).pack(side="left", padx=6, ipady=3)

        self._status_lbl = tk.Label(center, text="", font=SMALL, fg=ACCENT, bg=BG, anchor="w")
        self._status_lbl.pack(fill="x", pady=(5,5))

        self._log = tk.Text(center, bg=SURFACE, fg=MUTED, font=MONO, relief="flat", height=8, state="disabled")
        self._log.pack(fill="both", expand=True)

        # ── Rodapé (Botões de Fluxo) ──
        foot = tk.Frame(self, bg=SURFACE, padx=16, pady=16)
        foot.pack(fill="x", side="bottom")
        
        tk.Label(foot, text="Fluxo de Trabalho:", font=("Courier New", 8, "bold"), fg=MUTED, bg=SURFACE).pack(anchor="w", pady=(0,8))
        
        btn_frame = tk.Frame(foot, bg=SURFACE)
        btn_frame.pack(fill="x")
        
        self.btn1 = tk.Button(btn_frame, text="1. Extrair Pedidos", bg=ACCENT, fg=BG, font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=8, command=lambda: self._run_thread(self._task_extrair))
        self.btn1.pack(side="left", expand=True, fill="x", padx=(0,4))
        
        self.btn2 = tk.Button(btn_frame, text="2. Enriquecer / Excel", bg=ACCENT, fg=BG, font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=8, command=lambda: self._run_thread(self._task_enriquecer))
        self.btn2.pack(side="left", expand=True, fill="x", padx=4)
        
        self.btn3 = tk.Button(btn_frame, text="3. Classificar Final", bg=ACCENT, fg=BG, font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=8, command=lambda: self._run_thread(self._task_classificar))
        self.btn3.pack(side="left", expand=True, fill="x", padx=(4,0))

    def _on_config_saved(self, new_cfg):
        self.cfg = new_cfg
        auth_ml.save_config(self.cfg)
        self._set_status("Configurações salvas.", GREEN)

    def _start_oauth(self):
        if not self.cfg.get("client_id") or not self.cfg.get("redirect_uri"):
            messagebox.showwarning("Aviso", "Preencha Client ID e Redirect URI nas Configurações.")
            return
        webbrowser.open(f"https://auth.mercadolivre.com.br/authorization?response_type=code&client_id={self.cfg['client_id']}&redirect_uri={self.cfg['redirect_uri']}")
        self.after(1000, self._ask_code)

    def _ask_code(self):
        code = simpledialog.askstring("OAuth", "Cole o 'code' da URL:")
        if code:
            code = code.split("code=", 1)[-1].split("&")[0]
            try:
                self.auth = auth_ml.exchange_code_for_token(self.cfg, code)
                self.auth["seller_id"] = auth_ml.fetch_seller_id(self.auth["access_token"])
                auth_ml.save_auth(self.auth)
                self._set_status("Autenticado com sucesso!", GREEN)
            except Exception as e:
                self._log_write(f"Erro OAuth: {e}")

    def _revoke(self):
        auth_ml.clear_auth()
        self.auth = auth_ml.load_auth()
        self._set_status("Desconectado.", RED)

    def _set_status(self, msg, color=MUTED):
        self._status_lbl.config(text=msg, fg=color)

    def _log_write(self, msg):
        self._log.config(state="normal")
        self._log.insert("end", msg + "\n")
        self._log.see("end")
        self._log.config(state="disabled")
        self.update_idletasks()

    def _check_auth(self):
        if not self.auth.get("access_token"):
            messagebox.showwarning("Atenção", "Faça a autenticação primeiro!")
            return False
        return True

    def _run_thread(self, task_func):
        for b in (self.btn1, self.btn2, self.btn3): b.config(state="disabled")
        threading.Thread(target=self._worker_wrapper, args=(task_func,), daemon=True).start()

    def _worker_wrapper(self, task_func):
        try:
            task_func()
        except Exception as e:
            self.after(0, lambda: self._log_write(f"\n[ERRO] {e}"))
        finally:
            self.after(0, lambda: [b.config(state="normal") for b in (self.btn1, self.btn2, self.btn3)])

    # ── Tarefas ──
    def _task_extrair(self):
        if not self._check_auth(): return
        if not COOKIE_PATH.exists():
            raise RuntimeError(f"cookie.json não encontrado em {get_appdata_dir()}")
        self.after(0, lambda: self._log_write("\n--- PASSO 1: EXTRAÇÃO ---"))
        self.auth = core.executar_extracao(self.cfg, self.auth, self.v_inicio.get(), self.v_fim.get(), lambda m: self.after(0, self._log_write, m))
        auth_ml.save_auth(self.auth)

    def _task_enriquecer(self):
        if not self._check_auth(): return
        self.after(0, lambda: self._log_write("\n--- PASSO 2: ENRIQUECIMENTO ---"))
        self.auth = core.executar_enriquecimento(self.cfg, self.auth, lambda m: self.after(0, self._log_write, m))
        auth_ml.save_auth(self.auth)

    def _task_classificar(self):
        self.after(0, lambda: self._log_write("\n--- PASSO 3: CLASSIFICAÇÃO FINAL ---"))
        core.executar_classificacao(lambda m: self.after(0, self._log_write, m))


if __name__ == "__main__":
    App().mainloop()