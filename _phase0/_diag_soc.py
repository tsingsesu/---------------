import os, numpy as np, openpyxl
from scipy.optimize import linprog
from scipy.sparse import lil_matrix
W = r"C:\Users\lhl87\Desktop\高教社杯全国大学生数学建模大赛"
def ls(p, s=None):
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    ws = wb[s] if s else wb.worksheets[0]
    r = [list(x) for x in ws.iter_rows(values_only=True)]; wb.close(); return r
r1 = ls(os.path.join(W,"附件","附件1.xlsx"))
price = np.array([float(r[1]) for r in r1[1:]]); load = np.array([float(r[2]) for r in r1[1:]]); pv = np.array([float(r[3]) for r in r1[1:]])
K=144; DT=1/6; ETA=0.9; U=5000*DT; EMIN,EMAX,E0=1200.,10800.,6000.
NET=(load-pv)*DT
n=3*K
c=np.concatenate([price,np.zeros(2*K)])
A=lil_matrix((3*K,n)); b=np.zeros(3*K)
for k in range(K):
    A[k,k]=-1.; A[k,K+k]=1.; A[k,2*K+k]=-1.; b[k]=-NET[k]
for k in range(K):
    for j in range(k+1):
        A[K+k,K+j]=ETA; A[K+k,2*K+j]=-1/ETA
        A[K+K+k,K+j]=-ETA; A[K+K+k,2*K+j]=1/ETA
    b[K+k]=EMAX-E0; b[K+K+k]=E0-EMIN
Aeq=lil_matrix((1,n))
for j in range(K):
    Aeq[0,K+j]=ETA; Aeq[0,2*K+j]=-1/ETA
res=linprog(c,A_ub=A.tocsr(),b_ub=b,A_eq=Aeq.tocsr(),b_eq=np.array([0.]),bounds=[(0,None)]*K+[(0,U)]*K+[(0,U)]*K,method="highs")
x=res.x[:K]; u=res.x[K:2*K]; v=res.x[2*K:]
E=np.concatenate([[E0],E0+np.cumsum(ETA*u-v/ETA)])
print("cost=%.6f  qty=%.6f  sum_u=%.6f sum_v=%.6f" % (res.fun, x.sum(), u.sum(), v.sum()))
print("E min=%.6f  argmin=%d  E max=%.6f argmax=%d" % (E.min(), int(np.argmin(E)), E.max(), int(np.argmax(E))))
print("block boundary E (k=0,24,48,...,144):", np.round(E[::24],4))
print("E 前 30 个值:", np.round(E[:30],3))
print("E 中 60..90:", np.round(E[60:90],3))
print("E 后 30 个值:", np.round(E[-30:],3))
print("E<=1230 的索引:", np.where(E<=1230)[0])
print("E>=10770 的索引:", np.where(E>=10770)[0])
print("逐块 charge/discharge:")
for i in range(6):
    s=i*24
    print("   blk%d charge=%.4f discharge=%.4f dE=%.4f" % (i, u[s:s+24].sum(), v[s:s+24].sum(), ETA*u[s:s+24].sum()-v[s:s+24].sum()/ETA))
