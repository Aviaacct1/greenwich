"""Verification mirror v3 for the financing sheet group (increment 7 acceptance).
Re-implements the v3 Waterfall / Debt Schedule / Equity Returns grammar
independently in python: depreciation in the taxable base, maintenance capex
before external interest, growth capex after debt service, LTM from the last
actual year, forward reserve retained before the sweep. Iterates the circular
block to convergence and compares against the LibreOffice recalc (which is
known to under-iterate; Excel's 200-iteration setting converges like this
mirror). Author: Avia Solutions."""
import sys, re, openpyxl

RECALC = sys.argv[1]
wb = openpyxl.load_workbook(RECALC, data_only=True)
os_, fi, cxs = wb["Operations Summary"], wb["Financing Inputs"], wb["Capex"]
hdr = [wb["Waterfall"].cell(row=6, column=12+i).value for i in range(60)]
years = [h for h in hdr if isinstance(h, str) and re.match(r"^\d{4}[AF]$", h)]
Y0 = int(years[0][:4]); LA = max(int(y[:4]) for y in years if y.endswith("A"))
ENTRY, HOLD = LA+1, 10; EXIT = ENTRY+HOLD-1
jE, jX = ENTRY-Y0, EXIT-Y0
def col(j): return 12+j

D = {r: fi[f"D{r}"].value for r in (5,6,8,9,10,11,12,13,14,15,16,18,19,20,21,23,24,25,26,27,36)}
EB = {j: os_.cell(row=12, column=col(j)).value for j in range(jE, jX+1)}
RV = {j: os_.cell(row=10, column=col(j)).value for j in range(jE, jX+1)}
# Capex sheet rows by label (acyclic; cached values exact)
cxrow = {}
for row in cxs.iter_rows(min_col=4, max_col=4):
    v = row[0].value
    if v == "Total expansion (growth) capex": cxrow["gc"] = row[0].row
    if v == "Maintenance capex": cxrow["mt"] = row[0].row
    if v == "Total depreciation": cxrow["dt"] = row[0].row
GC = {j: cxs.cell(row=cxrow["gc"], column=col(j)).value or 0 for j in range(jE, jX+1)}
MT = {j: cxs.cell(row=cxrow["mt"], column=col(j)).value or 0 for j in range(jE, jX+1)}
DEP = {j: cxs.cell(row=cxrow["dt"], column=col(j)).value or 0 for j in range(jE, jX+1)}

LTM = os_.cell(row=12, column=col(LA-Y0)).value          # last actual year EBITDA
senior0 = D[8]*LTM; junior0 = D[9]*LTM
fees = D[13]*(senior0+junior0); txn = D[36]
dsr_fund = D[16]*(senior0*D[10]+junior0*D[11]+junior0*D[15])
uses = D[5]*LTM + fees + txn + dsr_fund
equity = uses - senior0 - junior0
gw = D[5]*LTM - D[21]; gw_life = D[20]

n = jX-jE+1
S = dict(so=[0.0]*n, sc=[0.0]*n, jo=[0.0]*n, jc=[0.0]*n, co=[0.0]*n, cc=[0.0]*n,
         div=[0.0]*n, spec=[0.0]*n, req=[0.0]*n)
for it in range(4000):
    prev = [tuple(S[k]) for k in S]
    for k in range(n):
        j = jE+k
        S["so"][k] = senior0 if k == 0 else S["sc"][k-1]
        S["jo"][k] = junior0 if k == 0 else S["jc"][k-1]
        S["co"][k] = dsr_fund if k == 0 else S["cc"][k-1]
        jrep = 0.0 if S["jo"][k] <= 0 else -min(junior0*D[15], S["jo"][k])
        S["jc"][k] = S["jo"][k] + jrep
        si = -((S["so"][k]+S["sc"][k])/2*D[10])
        ji = -((S["jo"][k]+S["jc"][k])/2*D[11])
        ext = si+ji
        gwa = -gw/gw_life if (Y0+j) < ENTRY+gw_life else 0.0
        dep = DEP[j]
        deb = RV[j]/365*D[23]; cred = RV[j]/365*D[24]
        wcm = (-deb+cred) if k == 0 else (-(deb-RV[j-1]/365*D[23])+(cred-RV[j-1]/365*D[24]))
        cint = (S["co"][k]+S["cc"][k])/2*D[12]
        taxable = EB[j]+dep+gwa+ext+cint + (-gwa if D[19] == "No" else 0.0)
        tax = 0.0 if taxable <= 0 else -taxable*D[18]
        cfo = EB[j]+wcm+cint+tax
        mcap = -MT[j]; gcap = -GC[j]
        if k < n-1:
            so_n = S["sc"][k]; jo_n = S["jc"][k]
            jrep_n = 0.0 if jo_n <= 0 else -min(junior0*D[15], jo_n)
            fwd = (so_n+S["sc"][k+1])/2*D[10] + (jo_n+S["jc"][k+1])/2*D[11] + (-jrep_n)
        else:
            fwd = 0.0
        req = fwd*D[16]; S["req"][k] = req
        surplus = S["co"][k]+cfo+mcap+ext-req
        sweep = 0.0 if surplus < 0 else -surplus*D[14]
        S["sc"][k] = max(S["so"][k]+sweep, 0.0)
        spaid = S["sc"][k]-S["so"][k]
        after = S["co"][k]+cfo+mcap+ext+spaid+jrep          # row 20
        after_g = after+gcap                                 # row 22
        ni = EB[j]+dep+gwa+ext+cint+tax
        S["div"][k] = -min(ni*D[25], after_g-req) if (ni > 0 and after_g-req > 0) else 0.0
        close_before = after_g+S["div"][k]
        cc = min(close_before, req) if D[26] == "Yes" else close_before
        S["spec"][k] = cc-close_before
        S["cc"][k] = cc
    if max(abs(a-b) for pk, kk in zip(prev, S.values()) for a, b in zip(pk, kk)) < 1e-11:
        break
print(f"mirror v3 converged after {it+1} sweeps | LTM ({LA}A): {round(LTM,1)}")
exit_ev = D[6]*EB[jX]
eq_exit = exit_ev - (S["sc"][n-1]+S["jc"][n-1]) + S["cc"][n-1] - exit_ev*D[27]
flows = [-(equity)-(S["div"][0]+S["spec"][0])] + [-(S["div"][k]+S["spec"][k]) for k in range(1, n)]
flows[-1] += eq_exit
def irr(cfs):
    lo, hi = -0.99, 10.0
    def npv(r): return sum(c/(1+r)**i for i, c in enumerate(cfs))
    for _ in range(200):
        mid = (lo+hi)/2
        if npv(mid) > 0: lo = mid
        else: hi = mid
    return (lo+hi)/2
print("mirror senior close:", [round(x, 2) for x in S["sc"]])
print("mirror cash close:  ", [round(x, 2) for x in S["cc"]])
print("mirror DSR checks (converged):", ["Ok" if S["cc"][k] >= S["req"][k]-1e-6 else "ERROR" for k in range(n)])
print("mirror equity IRR:", round(irr(flows), 5), "| MOIC:",
      round(sum(f for f in flows if f > 0)/equity, 4))
wf = wb["Waterfall"]; ds = wb["Debt Schedule"]; er = wb["Equity Returns"]
lo_cc = [wf.cell(row=28, column=col(jE+k)).value for k in range(n)]
print("LO deltas cash close:", [round(S["cc"][k]-lo_cc[k], 3) for k in range(n)])
print("LO equity IRR:", round(er["H15"].value, 5))
