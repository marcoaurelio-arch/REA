"""
Validação independente do modelo de viabilidade.

Recalcula em Python puro o fluxo de caixa dos 3 cenários a partir das
premissas lidas da planilha e compara com os valores calculados pelo Excel /
LibreOffice. Exige que a planilha tenha valores em cache (recalculada).

Uso:
    python scripts/validar_modelo.py [modelo.xlsx]
Saída: código 0 se todos os indicadores batem (tolerância 0,01%).
"""

import sys
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import gerar_modelo as gm  # noqa: E402  (reaproveita o mapa de linhas)


def irr(flows, lo=-0.99, hi=1.0, tol=1e-10):
    npv = lambda r: sum(f / (1 + r) ** t for t, f in enumerate(flows))
    for _ in range(300):
        mid = (lo + hi) / 2
        if npv(lo) * npv(mid) <= 0:
            hi = mid
        else:
            lo = mid
        if hi - lo < tol:
            break
    return (lo + hi) / 2


def simular(v, abl, aluguel, ac, ext):
    n = 13
    noi, re_, capex, venda, cvenda, ir = ([0.0] * n for _ in range(6))
    for t in range(n):
        op = max(0, t - 1)
        idx = (1 + v["reajuste"]) ** t
        occ = 0 if op == 0 else v["occ1"] if op == 1 else v["occ2"] if op == 2 else 1 - v["vacancia"]
        fcar = max(0, 12 - v["carencia"]) / 12 if op == 1 else 1
        pot = abl * aluguel * v["fator_alug"] * 12 * idx if op > 0 else 0
        rbe = pot * occ * fcar * (1 + v["outras"])
        re_[t] = rbe * (1 - v["inad"])
        vago = abl * (1 - occ) * v["custo_vago"] * 12 * idx if op > 0 else 0
        noi[t] = re_[t] * (1 - v["adm"] - v["nrep"] - v["reserva"]) - vago
        obra = 0
        if t == 0:
            obra = ac * v["obra_m2"] * v["p_obra_a0"]
            capex[t] = (v["terreno_rs"] * (1 + v["itbi"]) + v["projetos"] * ac * v["obra_m2"]
                        + v["taxas"] + obra * (1 + v["gerenc"] + v["conting"]))
        elif t == 1:
            obra = ac * v["obra_m2"] * (1 - v["p_obra_a0"]) * (1 + v["incc"])
            infra = ext * v["infra_m2"] * (1 + v["incc"])
            com = (v["comissao"] * abl * aluguel * v["fator_alug"] * (1 - v["vacancia"])
                   * (1 + v["reajuste"]))
            capex[t] = (obra * (1 + v["gerenc"] + v["conting"]) + infra * (1 + v["conting"])
                        + v["ti_m2"] * abl + com + v["marketing"])
    saida = int(v["ano_saida"])
    custo = sum(capex)
    noi_fwd = noi[saida + 1] if saida + 1 < n else noi[saida] * (1 + v["reajuste"])
    venda_v = noi_fwd / v["cap_saida"]
    cv = venda_v * v["custo_venda"]
    ir_v = max(0, venda_v - cv - custo) * v["ir_ganho"]
    fcl = []
    for t in range(n):
        if t > saida:
            fcl.append(0.0)
            continue
        f = noi[t] - re_[t] * v["tributos"] - capex[t]
        if t == saida:
            f += venda_v - cv - ir_v
        fcl.append(f)
    noi_e = noi[4] / (1 + v["reajuste"]) ** 4
    tma = v["tma"]
    return {
        "i_custo": custo,
        "i_noi": noi_e,
        "i_yoc": noi_e / custo,
        "i_valor": noi_e / v["cap_saida"],
        "i_tir": irr(fcl),
        "i_vpl": sum(f / (1 + tma) ** t for t, f in enumerate(fcl)),
        "i_moic": sum(f for f in fcl if f > 0) / -sum(f for f in fcl if f < 0),
    }


def main(path):
    # popula os mapas de linhas (P, M, F) rodando o gerador em memória
    from openpyxl import Workbook
    wb_tmp = Workbook()
    gm.build_instrucoes(wb_tmp)
    gm.build_premissas(wb_tmp)
    gm.build_mix(wb_tmp)
    gm.build_fluxo(wb_tmp)

    wb = load_workbook(path, data_only=True)
    pr, mx, fl = wb["Premissas"], wb["Mix"], wb["Fluxo"]
    val = lambda ref: mx[ref.replace("Mix!", "").replace("$", "")].value
    abl, aluguel, ac, ext = val(gm.M["abl"]), val(gm.M["aluguel"]), val(gm.M["ac"]), val(gm.M["ext"])
    if abl is None:
        print("ERRO: planilha sem valores em cache. Recalcule antes de validar.")
        return 2

    falhas = 0
    for col, nome in gm.CEN:
        v = {k: pr[f"{col}{r}"].value for k, r in gm.P.items()}
        esperado = simular(v, abl, aluguel, ac, ext)
        print(f"\n== {nome}")
        for k, e in esperado.items():
            obtido = fl[f"C{gm.F[(col, k)]}"].value
            ok = isinstance(obtido, (int, float)) and abs(obtido - e) <= max(1e-4 * abs(e), 1e-6)
            falhas += not ok
            print(f"  {'OK ' if ok else 'ERRO'} {k:10s} planilha={obtido!s:>22} python={e:,.6f}")
    print(f"\n{'VALIDAÇÃO OK' if falhas == 0 else f'{falhas} DIVERGÊNCIA(S)'}")
    return 1 if falhas else 0


if __name__ == "__main__":
    root = Path(__file__).resolve().parent.parent
    p = Path(sys.argv[1]) if len(sys.argv) > 1 else root / "modelo" / "Viabilidade_Strip_Mall.xlsx"
    sys.exit(main(p))
