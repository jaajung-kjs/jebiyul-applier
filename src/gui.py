"""tkinter 입력창과 입력 검증."""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from src.mapping import KINDS, SANAN_KIND_DEFAULT, default_sanan_kind
from src.lookup import SANAN_KINDS
from src.hwp_reader import read_hwp
from src import nomu_model


def validate_inputs(raw: dict) -> dict:
    """문자열 입력을 검증·형변환해 compute_rates용 params dict 반환.

    오류 시 ValueError(메시지) 발생.
    """
    def as_int(key, label):
        v = (raw.get(key) or "").replace(",", "").strip()
        if not v:
            raise ValueError(f"{label}을(를) 입력하세요.")
        try:
            return int(v)
        except ValueError:
            raise ValueError(f"{label}은(는) 숫자여야 합니다: {v!r}")

    kind = raw.get("kind")
    if kind not in KINDS:
        raise ValueError(f"공사종류를 선택하세요(허용: {KINDS}).")
    if raw.get("contract") not in ("경쟁", "수의"):
        raise ValueError("계약방법을 선택하세요(경쟁/수의).")
    if raw.get("sanjae_basis") not in ("한전", "조달청"):
        raise ValueError("산재 기준을 선택하세요(한전/조달청).")
    if not raw.get("jebiyul_path"):
        raise ValueError("제비율 파일을 선택하세요.")

    est_raw = (raw.get("est_cost") or "").replace(",", "").strip()
    if est_raw:
        try:
            est_cost = int(est_raw)
        except ValueError:
            raise ValueError(f"추정금액은 숫자여야 합니다: {est_raw!r}")
    else:
        est_cost = None

    sanan_kind = raw.get("sanan_kind") or default_sanan_kind(kind)
    if sanan_kind not in SANAN_KINDS:
        raise ValueError(f"산안비 공사종류를 선택하세요(허용: {list(SANAN_KINDS)}).")

    return {
        "jikjeop_cost": as_int("jikjeop_cost", "직접공사비"),
        "days": as_int("days", "공사기간(일)"),
        "kind": kind,
        "contract": raw["contract"],
        "sanjae_basis": raw["sanjae_basis"],
        "sanan_target": as_int("sanan_target", "산안비 대상액"),
        "jebiyul_path": raw["jebiyul_path"],
        "sanan_kind": sanan_kind,
        "est_cost": est_cost,
    }


PAD = 8          # 위젯 간 기본 여백
ENTRY_W = 22     # 입력칸 문자 폭
PATH_W = 46      # 파일명 라벨의 고정 폭(문자 단위) — 긴 경로가 창을 늘리지 못하게
NAME_MAX = 24    # 표시 글자 수 상한(한글은 폭이 넓어 라벨을 넘지 않게 보수적으로)


def _shorten(path, limit=NAME_MAX):
    """전체 경로 대신 파일명만, 길면 앞부분을 남기고 말줄임.

    라벨 폭을 고정(PATH_W)하고 글자 수도 제한해, 경로 길이가 창 크기를 좌우하지
    않게 한다. 파일 식별에 유리하도록 뒤가 아니라 앞을 남긴다.
    """
    if not path:
        return "선택된 파일 없음"
    name = os.path.basename(path)
    return name if len(name) <= limit else name[:limit - 1] + "…"


def _bind_wheel(canvas):
    """목록 위에서 마우스 휠 스크롤(플랫폼별 이벤트 차이 흡수)."""
    def on_wheel(e):
        if e.num == 5 or e.delta < 0:
            canvas.yview_scroll(1, "units")
        else:
            canvas.yview_scroll(-1, "units")

    def enter(_e):
        canvas.bind_all("<MouseWheel>", on_wheel)
        canvas.bind_all("<Button-4>", on_wheel)
        canvas.bind_all("<Button-5>", on_wheel)

    def leave(_e):
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            canvas.unbind_all(seq)

    canvas.bind("<Enter>", enter)
    canvas.bind("<Leave>", leave)


def run_app(on_submit):
    """tkinter 입력창을 띄우고 검증 후 on_submit(params)를 호출한다."""
    root = tk.Tk()
    root.title("적용근거 생성기")
    root.minsize(560, 640)
    try:
        root.call("tk", "scaling", 1.2)      # 고DPI에서 글자 뭉개짐 완화
    except tk.TclError:
        pass

    vars_ = {k: tk.StringVar() for k in
             ["jikjeop_cost", "days", "sanan_target", "est_cost", "jebiyul_path"]}
    kind = tk.StringVar(value=KINDS[1])
    contract = tk.StringVar(value="경쟁")
    sanjae = tk.StringVar(value="한전")
    sanan_kind = tk.StringVar(value=default_sanan_kind(KINDS[1]))
    hwp_path = tk.StringVar()
    jebiyul_label = tk.StringVar(value=_shorten(""))
    hwp_label = tk.StringVar(value=_shorten(""))
    count_label = tk.StringVar(value="hwp 파일을 선택하면 직종 목록이 표시됩니다.")
    nomu_vars: dict = {}

    outer = ttk.Frame(root, padding=PAD + 2)
    outer.pack(fill="both", expand=True)
    outer.columnconfigure(0, weight=1)
    outer.rowconfigure(2, weight=1)               # 직종 목록만 늘어남

    # ── 1. 입력 파일 ────────────────────────────────────────────────
    files = ttk.LabelFrame(outer, text=" 입력 파일 ", padding=PAD)
    files.grid(row=0, column=0, sticky="ew")
    files.columnconfigure(2, weight=1)

    def pick():
        p = filedialog.askopenfilename(
            title="제비율 파일 선택", filetypes=[("Excel 파일", "*.xlsx")])
        if p:
            vars_["jebiyul_path"].set(p)
            jebiyul_label.set(_shorten(p))

    ttk.Label(files, text="제비율").grid(row=0, column=0, sticky="w")
    ttk.Button(files, text="파일 선택…", command=pick, width=12).grid(
        row=0, column=1, padx=PAD)
    ttk.Label(files, textvariable=jebiyul_label, width=PATH_W,
              anchor="w", foreground="#444").grid(row=0, column=2, sticky="w")

    def pick_hwp():
        p = filedialog.askopenfilename(
            title="임금실태조사 파일 선택", filetypes=[("한글 파일", "*.hwp")])
        if not p:
            return
        for child in nomu_frame.winfo_children():
            child.destroy()
        nomu_vars.clear()
        try:
            report = read_hwp(p)
        except Exception as e:
            hwp_path.set("")
            hwp_label.set(_shorten(""))
            count_label.set("hwp 파일을 선택하면 직종 목록이 표시됩니다.")
            messagebox.showerror("hwp 읽기 실패", str(e))
            return
        hwp_path.set(p)
        hwp_label.set(_shorten(p))
        for name, on in nomu_model.preselect(list(report.order)):
            var = tk.BooleanVar(value=on)
            var.trace_add("write", lambda *_a: _refresh_count())
            nomu_vars[name] = var
            ttk.Checkbutton(nomu_frame, text=name, variable=var).pack(
                anchor="w", padx=4)
        _refresh_count()

    ttk.Label(files, text="노무임").grid(row=1, column=0, sticky="w", pady=(PAD, 0))
    ttk.Button(files, text="파일 선택…", command=pick_hwp, width=12).grid(
        row=1, column=1, padx=PAD, pady=(PAD, 0))
    ttk.Label(files, textvariable=hwp_label, width=PATH_W,
              anchor="w", foreground="#444").grid(
        row=1, column=2, sticky="w", pady=(PAD, 0))
    ttk.Label(files, text="(선택 — 넣으면 7.통신노무임 시트가 함께 생성됩니다)",
              foreground="#777").grid(row=2, column=0, columnspan=3,
                                      sticky="w", pady=(4, 0))

    # ── 2. 공사 정보 ────────────────────────────────────────────────
    info = ttk.LabelFrame(outer, text=" 공사 정보 ", padding=PAD)
    info.grid(row=1, column=0, sticky="ew", pady=(PAD + 2, 0))

    for i, (lab, key) in enumerate([("직접공사비(원)", "jikjeop_cost"),
                                    ("공사기간(일)", "days"),
                                    ("산안비 대상액(원)", "sanan_target"),
                                    ("추정금액(원)", "est_cost")]):
        ttk.Label(info, text=lab).grid(row=i, column=0, sticky="w",
                                       pady=2, padx=(0, PAD))
        ttk.Entry(info, textvariable=vars_[key], width=ENTRY_W,
                  justify="right").grid(row=i, column=1, sticky="w", pady=2)

    for i, (lab, var, vals) in enumerate([("공사종류", kind, list(KINDS)),
                                          ("계약방법", contract, ["경쟁", "수의"]),
                                          ("산재기준", sanjae, ["한전", "조달청"])]):
        ttk.Label(info, text=lab).grid(row=i, column=2, sticky="w",
                                       pady=2, padx=(PAD * 2, PAD))
        ttk.Combobox(info, textvariable=var, values=vals, state="readonly",
                     width=12).grid(row=i, column=3, sticky="w", pady=2)

    # 산안비 공사종류: 고용노동부 고시 분류(제비율 공사종류와 1:1 아님).
    # 공사종류를 바꾸면 기본값이 따라오되, 분리발주/부대공사에 따라 직접 바꿀 수 있다.
    ttk.Label(info, text="(선택 — 산안비 50억 이상 구간의 800억 기준 판정용)",
              foreground="#777").grid(row=3, column=2, columnspan=2, sticky="w",
                                      pady=2, padx=(PAD * 2, 0))

    ttk.Label(info, text="산안비 공사종류").grid(row=4, column=0, sticky="w",
                                          pady=(PAD, 2), padx=(0, PAD))
    ttk.Combobox(info, textvariable=sanan_kind, values=list(SANAN_KINDS),
                 state="readonly", width=14).grid(row=4, column=1, sticky="w",
                                                  pady=(PAD, 2))
    ttk.Label(info, text="(조경·전기·통신·소방을 분리발주·독립수행하면 특수건설공사)",
              foreground="#777").grid(row=4, column=2, columnspan=2,
                                      sticky="w", pady=(PAD, 2), padx=(PAD * 2, 0))
    kind.trace_add("write",
                   lambda *_a: sanan_kind.set(default_sanan_kind(kind.get())))

    # ── 3. 직종 선택 ────────────────────────────────────────────────
    nomu = ttk.LabelFrame(outer, text=" 직종 선택 ", padding=PAD)
    nomu.grid(row=2, column=0, sticky="nsew", pady=(PAD + 2, 0))
    nomu.columnconfigure(0, weight=1)
    nomu.rowconfigure(1, weight=1)

    bar = ttk.Frame(nomu)
    bar.grid(row=0, column=0, sticky="ew", pady=(0, PAD // 2))
    bar.columnconfigure(2, weight=1)

    def _set_all(on):
        for v in nomu_vars.values():
            v.set(on)

    ttk.Button(bar, text="전체 선택", width=10,
               command=lambda: _set_all(True)).grid(row=0, column=0)
    ttk.Button(bar, text="전체 해제", width=10,
               command=lambda: _set_all(False)).grid(row=0, column=1, padx=(4, 0))
    ttk.Label(bar, textvariable=count_label, foreground="#555").grid(
        row=0, column=2, sticky="e")

    # 목록: 캔버스 + 스크롤바를 같은 프레임에 붙여 항상 맞닿게 한다
    box = ttk.Frame(nomu, relief="solid", borderwidth=1)
    box.grid(row=1, column=0, sticky="nsew")
    nomu_canvas = tk.Canvas(box, height=220, highlightthickness=0,
                            background="white")
    nomu_scroll = ttk.Scrollbar(box, orient="vertical",
                                command=nomu_canvas.yview)
    nomu_scroll.pack(side="right", fill="y")
    nomu_canvas.pack(side="left", fill="both", expand=True)
    nomu_canvas.configure(yscrollcommand=nomu_scroll.set)

    nomu_frame = ttk.Frame(nomu_canvas)
    win = nomu_canvas.create_window((0, 0), window=nomu_frame, anchor="nw")
    nomu_frame.bind("<Configure>", lambda e: nomu_canvas.configure(
        scrollregion=nomu_canvas.bbox("all")))
    # 내부 프레임 폭을 캔버스 폭에 맞춰 좌측에 몰리지 않게
    nomu_canvas.bind("<Configure>",
                     lambda e: nomu_canvas.itemconfigure(win, width=e.width))
    _bind_wheel(nomu_canvas)

    def _refresh_count():
        if not nomu_vars:
            count_label.set("hwp 파일을 선택하면 직종 목록이 표시됩니다.")
            return
        n = sum(1 for v in nomu_vars.values() if v.get())
        count_label.set(f"{n} / {len(nomu_vars)} 직종 선택됨")

    # ── 4. 실행 ────────────────────────────────────────────────────
    def submit():
        raw = {k: v.get() for k, v in vars_.items()}
        raw.update(kind=kind.get(), contract=contract.get(),
                   sanjae_basis=sanjae.get(), sanan_kind=sanan_kind.get())
        try:
            params = validate_inputs(raw)
        except ValueError as e:
            messagebox.showerror("입력 오류", str(e))
            return
        params["hwp_path"] = hwp_path.get() or None
        params["selected_nomu"] = [n for n, v in nomu_vars.items() if v.get()] or None
        run_btn.state(["disabled"])
        root.update_idletasks()
        try:
            out = on_submit(params)
            if params.get("selected_nomu"):
                messagebox.showinfo("완료", f"생성 완료 (7.통신노무임 포함)\n\n{out}")
            else:
                messagebox.showinfo("완료", f"생성 완료\n\n{out}")
        except Exception as e:
            messagebox.showerror("생성 실패", str(e))
        finally:
            run_btn.state(["!disabled"])

    run = ttk.Frame(outer)
    run.grid(row=3, column=0, sticky="ew", pady=(PAD + 2, 0))
    run.columnconfigure(0, weight=1)
    run_btn = ttk.Button(run, text="적용근거 생성", command=submit)
    run_btn.grid(row=0, column=0, sticky="e", ipadx=PAD * 2, ipady=2)

    root.mainloop()
