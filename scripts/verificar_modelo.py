# Recalcula el modelo en Python puro para verificar coherencia de resultados.
N=24
area=3500; unid=20; precio=560000
ventas=unid*precio
ini_pct=0.10; saldo_pct=0.90
terreno=2500000; cm2=1650; constr=area*cm2
lic=constr*0.06; com=ventas*0.03; adm=ventas*0.04
pfin=0.60; linea=constr*pfin
tea=0.12; tem=(1+tea)**(1/12)-1
cok=0.15; cokm=(1+cok)**(1/12)-1
ir=0.295

# cronograma pesos
scurve={4:2,5:3,6:4,7:5,8:6,9:7,10:8,11:8,12:8,13:8,14:8,15:7,16:7,17:6,18:5,19:4,20:3,21:2}
w=[scurve.get(i,0) for i in range(1,N+1)]; sw=sum(w)
pct=[x/sw for x in w]
val=[p*constr for p in pct]
des=[v*pfin for v in val]
tot_des=sum(des)

# ingresos
recw=[1 if 2<=i<=20 else 0 for i in range(1,N+1)]; srec=sum(recw)
entw=[{22:1,23:2,24:2}.get(i,0) for i in range(1,N+1)]; sent=sum(entw)
cuota_ini=[(recw[i]/srec)*ventas*ini_pct for i in range(N)]
cobro=[(entw[i]/sent)*ventas*saldo_pct for i in range(N)]

# egresos (negativos)
terr=[-terreno if i==0 else 0 for i in range(N)]
con=[-val[i] for i in range(N)]
licr=[-lic/3 if i in (0,1,2) else 0 for i in range(N)]
comr=[-com/20 if 2<=i<=21 else 0 for i in range(N)]
admr=[-adm/N for i in range(N)]
teg=[terr[i]+con[i]+licr[i]+comr[i]+admr[i] for i in range(N)]

# deuda
sini=[0]*N; sfin=[0]*N; amort=[0]*N; intr=[0]*N
for i in range(N):
    sini[i]= sfin[i-1] if i>0 else 0
    desL=des[i]
    amort[i]=-(entw[i]/sent)*tot_des
    intr[i]=-sini[i]*tem
    sfin[i]=sini[i]+desL+amort[i]

eco=[cuota_ini[i]+cobro[i]+teg[i] for i in range(N)]
fin=[eco[i]+des[i]+amort[i]+intr[i] for i in range(N)]

def npv(rate,flows):
    return sum(f/(1+rate)**(k+1) for k,f in enumerate(flows))
def irr(flows):
    lo,hi=-0.99,10.0
    for _ in range(200):
        mid=(lo+hi)/2
        v=sum(f/(1+mid)**k for k,f in enumerate(flows))
        if v>0: lo=mid
        else: hi=mid
    return (lo+hi)/2

print(f"Ventas totales:        S/ {ventas:,.0f}")
print(f"Costo construccion:    S/ {constr:,.0f}")
print(f"Inversion total:       S/ {terreno+constr+lic+com+adm:,.0f}")
print(f"Linea max banco:       S/ {linea:,.0f}")
print(f"Total desembolsado:    S/ {tot_des:,.0f}")
print(f"Suma valorizaciones:   S/ {sum(val):,.0f}  (=costo constr? {abs(sum(val)-constr)<1})")
print(f"Saldo final linea m24: S/ {sfin[-1]:,.2f}  (~0? {abs(sfin[-1])<1})")
print(f"Total intereses:       S/ {sum(intr):,.0f}")
print(f"Total ingresos:        S/ {sum(cuota_ini)+sum(cobro)+tot_des:,.0f}")
uai=ventas-terreno-constr-lic-com-adm+sum(intr)
impuesto=-max(uai,0)*ir
neta=uai+impuesto
print(f"Utilidad bruta:        S/ {ventas-terreno-constr-lic:,.0f}")
print(f"Utilidad antes imp:    S/ {uai:,.0f}")
print(f"Utilidad neta:         S/ {neta:,.0f}  margen {neta/ventas*100:.1f}%")
print(f"VAN economico:         S/ {npv(cokm,eco):,.0f}")
print(f"TIR economica mensual: {irr(eco)*100:.2f}%  anual {((1+irr(eco))**12-1)*100:.1f}%")
print(f"VAN financiero:        S/ {npv(cokm,fin):,.0f}")
print(f"TIR financiera mensual:{irr(fin)*100:.2f}%  anual {((1+irr(fin))**12-1)*100:.1f}%")
print(f"Flujo fin. acum final: S/ {sum(fin):,.0f}")
