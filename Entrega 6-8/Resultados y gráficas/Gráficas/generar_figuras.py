"""
Genera las figuras del proyecto "Cafeterías de especialidad dentro y fuera del Park Way".

Uso (desde la carpeta del repositorio):
    pip install openpyxl matplotlib
    python codigo/generar_figuras.py

Lee:   Base_de_datos_conciliada_ParkWay.xlsx   (hojas Parametros, Locales y Encuesta_valida)
Crea:  figuras/fig1_mapa_dentro_fuera.png ... fig5_llegada_a_proposito.png

Reglas:
- Un local es "Dentro" si su distancia al eje del corredor (Carrera 24, calles 36 a 45)
  es menor o igual al umbral de la hoja Parametros (celda B3, en metros); si no, "Fuera".
- Solo entran al análisis los locales con Estado de la visita = "Respondió".
- Dos colores fijos en todas las figuras: azul = dentro, naranja = fuera.
"""
from pathlib import Path
import openpyxl
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D

RAIZ = Path(__file__).resolve().parent.parent
BASE = RAIZ / 'Base_de_datos_conciliada_ParkWay.xlsx'
SALIDA = RAIZ / 'figuras'
SALIDA.mkdir(exist_ok=True)

C_IN, C_OUT = '#2a78d6', '#eb6834'
INK, INK2, GRID = '#0b0b0b', '#52514e', '#e3e2de'
plt.rcParams.update({'font.family': 'DejaVu Sans', 'text.color': INK, 'axes.labelcolor': INK2,
                     'xtick.color': INK2, 'ytick.color': INK2, 'axes.edgecolor': GRID})

# ------------------------------------------------------------------ datos
wb = openpyxl.load_workbook(BASE, data_only=True)
umbral = wb['Parametros']['B3'].value

locales = {}
for r in list(wb['Locales'].iter_rows(values_only=True))[1:]:
    if not r[0]:
        continue
    _id, nombre, _ig, lat, lon, _dir, dist, _zona, estado = r[:9]
    zona = ('Dentro' if dist <= umbral else 'Fuera') if isinstance(dist, (int, float)) else None
    locales[nombre] = dict(lat=lat, lon=lon, zona=zona, estado=estado)

filas = list(wb['Encuesta_valida'].iter_rows(values_only=True))
enc = [c.strip() if c else c for c in filas[0]]


def col(prefijo):
    return next(i for i, h in enumerate(enc) if h and h.startswith(prefijo))


C_NOMBRE = col('Nombre del c')
C_CLIENTES = col('¿Quiénes son sus clientes')
C_PROPOSITO = col('¿Qué parte de sus clientes')
C_SEMANA = col('¿Cuál es la zona horaria más frecuente de visita de los clientes entre')
C_PERM = col('Permanencia')

datos = [r for r in filas[1:] if r[C_NOMBRE]]
for r in datos:
    assert locales[r[C_NOMBRE]]['zona'] is not None, r[C_NOMBRE]
dentro = [r for r in datos if locales[r[C_NOMBRE]]['zona'] == 'Dentro']
fuera = [r for r in datos if locales[r[C_NOMBRE]]['zona'] == 'Fuera']
N_IN, N_OUT = len(dentro), len(fuera)
LBL_IN = f'Dentro del Park Way (n = {N_IN})'
LBL_OUT = f'Fuera del Park Way (n = {N_OUT})'
NOTA = 'Respuestas de los encargados; son estimaciones del personal, no mediciones.'


# ------------------------------------------------------------------ barras agrupadas
def barras(archivo, titulo, subtitulo, categorias, c, multiple=False, nota=NOTA):
    def contar(grupo, cat):
        if multiple:
            return sum(1 for r in grupo if cat in (r[c] or ''))
        return sum(1 for r in grupo if r[c] == cat)

    fig, ax = plt.subplots(figsize=(8.2, 0.62 * len(categorias) + 2.2), dpi=200)
    h = 0.36
    for i, (clave, corto) in enumerate(categorias):
        y = len(categorias) - 1 - i
        for desp, grupo, n, color in ((+h / 2 + 0.01, dentro, N_IN, C_IN), (-h / 2 - 0.01, fuera, N_OUT, C_OUT)):
            k = contar(grupo, clave)
            pct = 100 * k / n
            ax.barh(y + desp, pct, height=h, color=color, edgecolor='white', linewidth=1)
            ax.text(pct + 1.5, y + desp, f'{k} de {n}', va='center', fontsize=8.5, color=INK2)
    ax.set_yticks(range(len(categorias)))
    ax.set_yticklabels([c_ for _, c_ in categorias][::-1], fontsize=9.5, color=INK)
    ax.set_xlim(0, 112)
    ax.set_xticks([0, 25, 50, 75, 100])
    ax.set_xticklabels(['0 %', '25 %', '50 %', '75 %', '100 %'], fontsize=8.5)
    ax.xaxis.grid(True, color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    for s in ('top', 'right', 'left'):
        ax.spines[s].set_visible(False)
    ax.tick_params(length=0)
    ax.set_xlabel('Porcentaje de locales de cada zona', fontsize=9)
    fig.suptitle(titulo, x=0.01, ha='left', fontsize=12.5, fontweight='bold', color=INK)
    fig.text(0.01, 0.925 if len(categorias) > 3 else 0.9, subtitulo, ha='left', fontsize=9, color=INK2)
    ax.legend(handles=[Line2D([0], [0], marker='s', color='w', markerfacecolor=C_IN, markersize=9, label=LBL_IN),
                       Line2D([0], [0], marker='s', color='w', markerfacecolor=C_OUT, markersize=9, label=LBL_OUT)],
              loc='lower right', frameon=False, fontsize=8.5)
    fig.text(0.01, 0.01, nota, ha='left', fontsize=7.5, color=INK2)
    fig.tight_layout(rect=(0, 0.04, 1, 0.9))
    fig.savefig(SALIDA / archivo, facecolor='white')
    plt.close(fig)


barras('fig2_tipo_de_cliente.png', 'Tipo de cliente frecuente',
       'Porcentaje de locales cuyo encargado menciona cada tipo de cliente (podía marcar varios)',
       [('Trabajadores de oficinas cercanas', 'Trabajadores de oficinas'), ('Vecinos del barrio', 'Vecinos del barrio'),
        ('Personas de otras zonas que vienen a propósito a la cafetería', 'Vienen a propósito de otras zonas'),
        ('Pacientes, acompañantes o personal de salud', 'Pacientes y personal de salud'), ('Turistas', 'Turistas')],
       C_CLIENTES, multiple=True, nota=NOTA + ' Estudiantes o profesores: ver Anexo B.')
barras('fig3_horario_entre_semana.png', 'Franja de mayor afluencia entre semana',
       'Porcentaje de locales según la franja que su encargado señala como la más frecuente',
       [('9-12pm', '9 a. m. a 12 m.'), ('12-2pm', '12 m. a 2 p. m.'), ('2-5pm', '2 a 5 p. m.'), ('5-8pm', '5 a 8 p. m.')],
       C_SEMANA)
barras('fig4_permanencia.png', 'Permanencia típica de los clientes',
       'Porcentaje de locales según el tiempo que, según su encargado, se queda un cliente',
       [('Menos de 30 minutos', 'Menos de 30 min'), ('De 30-60 minutos', '30 a 60 min'),
        ('De 1-2 horas', '1 a 2 horas'), ('Más de 2 horas', 'Más de 2 horas')], C_PERM)
barras('fig5_llegada_a_proposito.png', 'Clientes que llegan a propósito desde otras zonas',
       'Porcentaje de locales según la parte de sus clientes que, según el encargado, vienen a propósito',
       [('Del 0 al 25%', '0 a 25 %'), ('Del 25 al 50%', '25 a 50 %'), ('Del 50 al 75%', '50 a 75 %'),
        ('Más de 75%', 'Más de 75 %'), ('No sabe no responde', 'No sabe / no responde')], C_PROPOSITO)

# ------------------------------------------------------------------ mapa
fig, ax = plt.subplots(figsize=(6.4, 7.4), dpi=200)
for nombre, d in locales.items():
    if d['estado'] != 'Respondió':
        ax.scatter(d['lon'], d['lat'], s=70, facecolor='white', edgecolor=INK2, linewidth=1.2, zorder=3)
    else:
        ax.scatter(d['lon'], d['lat'], s=95, color=C_IN if d['zona'] == 'Dentro' else C_OUT,
                   edgecolor='white', linewidth=1.5, zorder=3)
OFF = {'Le tiende': (-8, -3, 'right'), 'Quipile': (8, 2, 'left'), 'La cesta': (8, 0, 'left'), 'Marttina Cafe': (8, -3, 'left'),
       'Somos': (-8, 0, 'right'), 'Coffelia': (8, 2, 'left'), 'Dulce Herencia': (-8, -4, 'right'),
       'Bizcotela cafe restaurante': (-8, 0, 'right'), 'Hora de café': (-8, -6, 'right'), 'Mistral': (8, 3, 'left'),
       'Rolo café': (8, -3, 'left'), 'Groso Café Argentino': (8, -6, 'left'), 'Doppio Café': (8, 5, 'left')}
CORTO = {'Bizcotela cafe restaurante': 'Bizcotela', 'Groso Café Argentino': 'Groso', 'Betty la fea, enamorados café': 'Betty la fea',
         'Bogota Coffee Roasters': 'Bogotá Coffee Roasters', 'Beraka Café Galeria': 'Beraka', 'Marttina Cafe': 'Marttina',
         'Doppio Café': 'Doppio (cerrado, excluido)'}
for nombre, d in locales.items():
    dx, dy, ha = OFF.get(nombre, (8, 0, 'left'))
    ax.annotate(CORTO.get(nombre, nombre), (d['lon'], d['lat']), xytext=(dx, dy), textcoords='offset points',
                ha=ha, va='center', fontsize=7.5, color=INK2 if d['estado'] != 'Respondió' else INK)
ax.set_xlim(-74.0815, -74.0690)
ax.set_ylim(4.6235, 4.6440)
ax.set_xticks([]); ax.set_yticks([])
for s in ax.spines.values():
    s.set_color(GRID)
ax.set_aspect(1)
ax.set_title('Locales visitados: dentro y fuera del Park Way', loc='left', fontsize=11.5, fontweight='bold', color=INK)
ax.legend(handles=[Line2D([0], [0], marker='o', color='w', markerfacecolor=C_IN, markersize=10, label=LBL_IN),
                   Line2D([0], [0], marker='o', color='w', markerfacecolor=C_OUT, markersize=10, label=LBL_OUT),
                   Line2D([0], [0], marker='o', color='w', markerfacecolor='white', markeredgecolor=INK2, markersize=9,
                          label='Excluido: cerrado permanentemente (1)')],
          loc='upper left', frameon=True, framealpha=0.95, edgecolor=GRID, fontsize=8)
fig.text(0.01, 0.01, f'{len(datos)} locales válidos de {len(locales)} en la base. Norte hacia arriba; escala aproximada.',
         fontsize=7.5, color=INK2)
fig.tight_layout(rect=(0, 0.02, 1, 1))
fig.savefig(SALIDA / 'fig1_mapa_dentro_fuera.png', facecolor='white')
plt.close(fig)

print(f'Listo: {N_IN} locales dentro y {N_OUT} fuera. Figuras en {SALIDA}')
