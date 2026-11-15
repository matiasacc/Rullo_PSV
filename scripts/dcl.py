import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
fig, axs = plt.subplots(2,1,figsize=(8,5.6), dpi=130)
def beam(ax, title, lab, wfull):
    ax.set_xlim(-1.4,11.2); ax.set_ylim(-2.4,2.9); ax.axis("off")
    ax.text(-1.3,2.6,title,fontsize=10,weight="bold")
    ax.plot([0,10],[0,0],color="k",lw=3)
    for x0,n in ((0,"1"),(10,"2")):
        ax.add_patch(Polygon([[x0,0],[x0-0.35,-0.55],[x0+0.35,-0.55]],closed=True,fill=False,lw=1.2))
        ax.annotate("",xy=(x0,-0.6),xytext=(x0,-1.4),arrowprops=dict(arrowstyle="->",lw=1.2))
        ax.text(x0,-1.75,"R"+n+lab,ha="center",fontsize=10)
    a,e=4.0,8.6
    ax.plot([a,e],[1.1,1.1],color="k",lw=0.8)
    for i in range(8):
        xi=a+i*(e-a)/7
        ax.annotate("",xy=(xi,0.08),xytext=(xi,1.1),arrowprops=dict(arrowstyle="->",lw=0.9))
    ax.text((a+e)/2,1.3,"q"+lab+" en B = n · ancho",ha="center",fontsize=9)
    if wfull: ax.text(5,-0.45,"pesos propios por tramo: muñones, tubo y engomado (opcional)",ha="center",fontsize=8,style="italic")
    def dim(x1,x2,y,t):
        ax.annotate("",xy=(x1,y),xytext=(x2,y),arrowprops=dict(arrowstyle="<->",lw=0.8))
        ax.text((x1+x2)/2,y+0.12,t,ha="center",fontsize=9)
    dim(0,a,1.9,"a = x_tope - B"); dim(a,e,1.9,"B"); dim(e,10,1.9,"b")
    ax.plot([e,e],[0.1,1.5],color="k",lw=2); ax.text(e+0.1,1.25,"tope",fontsize=8)
    ax.add_patch(Polygon([[0,-0.08],[1.0,-0.08],[1.0,0.08],[0,0.08]],closed=True,fc="0.8",ec="k",lw=0.6))
    ax.add_patch(Polygon([[9.0,-0.08],[10,-0.08],[10,0.08],[9.0,0.08]],closed=True,fc="0.8",ec="k",lw=0.6))
    dim(0,10,-2.25,"L (entre centros de rodamientos)")
    ax.text(-1.2,0.1,"x →",fontsize=9)
beam(axs[0],"Plano y: tensión del manchón  (qy = 2·n·T / B)","y",False)
beam(axs[1],"Plano z: gravedad  (peso propio + ½ peso del manchón)","z",True)
axs[0].annotate("",xy=(10.8,0.35),xytext=(10.8,-0.35),arrowprops=dict(arrowstyle="->",connectionstyle="arc3,rad=1.2",lw=1.2))
axs[0].text(10.2,0.6,"T (motor)",fontsize=9)
plt.tight_layout(); plt.savefig("dcl.png")
