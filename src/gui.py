"""tkinter 입력창과 입력 검증."""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from src.mapping import KINDS
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

    return {
        "jikjeop_cost": as_int("jikjeop_cost", "직접공사비"),
        "days": as_int("days", "공사기간(일)"),
        "kind": kind,
        "contract": raw["contract"],
        "sanjae_basis": raw["sanjae_basis"],
        "sanan_target": as_int("sanan_target", "산안비 대상액"),
        "jebiyul_path": raw["jebiyul_path"],
    }


def run_app(on_submit):
    """tkinter 입력창을 띄우고 검증 후 on_submit(params)를 호출한다."""
    root = tk.Tk()
    root.title("적용근거 생성기")

    vars_ = {k: tk.StringVar() for k in
             ["jikjeop_cost", "days", "sanan_target", "jebiyul_path"]}
    kind = tk.StringVar(value=KINDS[1])
    contract = tk.StringVar(value="경쟁")
    sanjae = tk.StringVar(value="한전")
    hwp_path = tk.StringVar()
    nomu_vars: dict = {}

    rows = [
        ("직접공사비(원)", "jikjeop_cost"),
        ("공사기간(일)", "days"),
        ("산안비 대상액(원)", "sanan_target"),
    ]
    for i, (lab, key) in enumerate(rows):
        ttk.Label(root, text=lab).grid(row=i, column=0, sticky="e")
        ttk.Entry(root, textvariable=vars_[key]).grid(row=i, column=1)

    ttk.Label(root, text="공사종류").grid(row=3, column=0, sticky="e")
    ttk.Combobox(root, textvariable=kind, values=KINDS,
                 state="readonly").grid(row=3, column=1)

    ttk.Label(root, text="계약방법").grid(row=4, column=0, sticky="e")
    ttk.Combobox(root, textvariable=contract, values=["경쟁", "수의"],
                 state="readonly").grid(row=4, column=1)

    ttk.Label(root, text="산재기준").grid(row=5, column=0, sticky="e")
    ttk.Combobox(root, textvariable=sanjae, values=["한전", "조달청"],
                 state="readonly").grid(row=5, column=1)

    def pick():
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if p:
            vars_["jebiyul_path"].set(p)

    ttk.Button(root, text="제비율 파일 선택",
               command=pick).grid(row=6, column=0, columnspan=2)
    ttk.Label(root, textvariable=vars_["jebiyul_path"]).grid(
        row=7, column=0, columnspan=2)

    # --- 임금실태조사(hwp) 선택 + 직종 체크리스트 (선택 사항) ---
    nomu_canvas = tk.Canvas(root, height=150, highlightthickness=0)
    nomu_scroll = ttk.Scrollbar(root, orient="vertical",
                                command=nomu_canvas.yview)
    nomu_frame = ttk.Frame(nomu_canvas)
    nomu_frame.bind(
        "<Configure>",
        lambda e: nomu_canvas.configure(scrollregion=nomu_canvas.bbox("all")),
    )
    nomu_canvas.create_window((0, 0), window=nomu_frame, anchor="nw")
    nomu_canvas.configure(yscrollcommand=nomu_scroll.set)

    def pick_hwp():
        p = filedialog.askopenfilename(filetypes=[("HWP", "*.hwp")])
        if not p:
            return
        hwp_path.set(p)
        for child in nomu_frame.winfo_children():
            child.destroy()
        nomu_vars.clear()
        try:
            report = read_hwp(p)
        except Exception as e:
            hwp_path.set("")
            messagebox.showerror("hwp 읽기 실패", str(e))
            return
        for name, on in nomu_model.preselect(list(report.order)):
            var = tk.BooleanVar(value=on)
            nomu_vars[name] = var
            tk.Checkbutton(nomu_frame, text=name, variable=var).pack(anchor="w")

    ttk.Button(root, text="임금실태조사(hwp) 파일 선택",
               command=pick_hwp).grid(row=8, column=0, columnspan=2)
    ttk.Label(root, textvariable=hwp_path).grid(
        row=9, column=0, columnspan=2)
    nomu_canvas.grid(row=10, column=0, sticky="nsew")
    nomu_scroll.grid(row=10, column=1, sticky="ns")

    def submit():
        raw = {k: v.get() for k, v in vars_.items()}
        raw.update(kind=kind.get(), contract=contract.get(),
                   sanjae_basis=sanjae.get())
        try:
            params = validate_inputs(raw)
        except ValueError as e:
            messagebox.showerror("입력 오류", str(e))
            return
        params["hwp_path"] = hwp_path.get() or None
        params["selected_nomu"] = [n for n, v in nomu_vars.items() if v.get()] or None
        try:
            out = on_submit(params)
            if params.get("selected_nomu"):
                messagebox.showinfo("완료", f"생성 완료 (7.통신노무임 포함):\n{out}")
            else:
                messagebox.showinfo("완료", f"생성 완료:\n{out}")
        except Exception as e:
            messagebox.showerror("생성 실패", str(e))

    ttk.Button(root, text="적용근거 생성",
               command=submit).grid(row=11, column=0, columnspan=2)
    root.mainloop()
