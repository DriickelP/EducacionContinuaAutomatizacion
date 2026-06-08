import flet as ft
from google.oauth2 import service_account
from googleapiclient.discovery import build

# --- CONFIGURACIÓN DE CREDENCIALES ---
CREDENTIALS_FILE = "educacion-continua-490016-7ff6875c1d8f.json" # Asegúrate de que este archivo esté en la misma carpeta
ID_SEGUIMIENTO_NUBE = "1szQLci5kxO10bTGyzQ1wQrQ9z_hITA0uAUHEZ-loaJU"

def obtener_datos_nube():
    """Conecta a Google Sheets y extrae todas las filas de TEC y ESP."""
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    serv = build("sheets", "v4", credentials=creds, cache_discovery=False)
    
    res_tec = serv.spreadsheets().values().get(spreadsheetId=ID_SEGUIMIENTO_NUBE, range="'SEGUIMIENTO TEC'!A:K").execute()
    res_esp = serv.spreadsheets().values().get(spreadsheetId=ID_SEGUIMIENTO_NUBE, range="'SEGUIMIENTO ESP'!A:M").execute()
    
    return res_tec.get("values", []), res_esp.get("values", [])

def main(page: ft.Page):
    page.title = "Dashboard FIME EC"
    page.theme_mode = ft.ThemeMode.DARK
    page.padding = 20
    page.window.width = 400
    page.window.height = 800

    # Pantalla de carga
    loading_ring = ft.ProgressRing()
    loading_text = ft.Text("Conectando con Google Sheets...", size=16)
    page.add(ft.Column([loading_ring, loading_text], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER))
    
    # Descargar datos
    datos_tec, datos_esp = obtener_datos_nube()
    page.clean() # Quitamos la pantalla de carga

    # Función auxiliar para extraer periodos únicos (columna 0)
    def obtener_periodos(datos):
        periodos = set()
        for fila in datos[1:]: # Omitir encabezados
            if len(fila) > 0 and fila[0].strip():
                try:
                    periodos.add(str(int(float(fila[0]))))
                except ValueError:
                    pass
        return sorted(list(periodos), reverse=True)

    periodos_tec = obtener_periodos(datos_tec)
    periodos_esp = obtener_periodos(datos_esp)

    # ==========================================
    # COMPONENTES VISUALES (TARJETAS)
    # ==========================================
    def crear_tarjeta_metrica(titulo, control_texto, color_fondo):
        return ft.Container(
            content=ft.Column([
                ft.Text(titulo, size=12, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE_70),
                control_texto, # Pinta el control directamente, sin meterlo en otro ft.Text
            ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            bgcolor=color_fondo,
            padding=15,
            border_radius=10,
            expand=True
        )

    # ==========================================
    # PESTAÑA TECNICA (TEC)
    # ==========================================
    dd_periodo_tec = ft.Dropdown(
        label="Selecciona el Periodo",
        options=[ft.dropdown.Option(p) for p in periodos_tec],
        value=periodos_tec[0] if periodos_tec else None,
        width=200
    )

    # Contenedores para las métricas TEC
    tec_tot_alumnos = ft.Text("0", size=28, weight=ft.FontWeight.BOLD, color="#2ed573")
    tec_form_contestado = ft.Text("0", size=28, weight=ft.FontWeight.BOLD, color="#2ed573")
    tec_papeleria = ft.Text("0", size=28, weight=ft.FontWeight.BOLD, color="#2ed573")
    tec_pendientes = ft.Text("0", size=28, weight=ft.FontWeight.BOLD, color="#ff4757")
    tec_avance = ft.Text("0%", size=28, weight=ft.FontWeight.BOLD, color="#2ed573")

    def actualizar_dashboard_tec(e=None):
        periodo_sel = dd_periodo_tec.value
        total = 0; contestaron = 0; papeleria = 0; pendientes = 0
        
        for fila in datos_tec[1:]:
            if len(fila) > 0 and str(fila[0]).startswith(str(periodo_sel)):
                total += 1
                if len(fila) > 5 and fila[5].strip().upper() == "SI":
                    contestaron += 1
                if len(fila) > 7 and fila[7].strip().upper() == "SI":
                    papeleria += 1
                if len(fila) > 8 and fila[8].strip() != "Si":
                    pendientes += 1

        avance = (contestaron / total * 100) if total > 0 else 0

        tec_tot_alumnos.value = str(total)
        tec_form_contestado.value = str(contestaron)
        tec_papeleria.value = str(papeleria)
        tec_pendientes.value = str(pendientes)
        tec_avance.value = f"{avance:.2f}%"
        page.update()

    dd_periodo_tec.on_change = actualizar_dashboard_tec

    vista_tec = ft.Column([
        ft.Row([dd_periodo_tec], alignment=ft.MainAxisAlignment.CENTER),
        ft.Row([
            crear_tarjeta_metrica("Pendientes", tec_pendientes, "#1e272e"),
            crear_tarjeta_metrica("% Avance", tec_avance, "#1e272e"),
        ]),
        ft.Divider(height=20, color=ft.Colors.TRANSPARENT),
        ft.Text("RESUMEN GENERAL", size=16, weight="bold"),
        ft.Row([
            crear_tarjeta_metrica("Total Alumnos", tec_tot_alumnos, "#2d3436"),
            crear_tarjeta_metrica("Contestados", tec_form_contestado, "#2d3436"),
        ]),
        ft.Row([
            crear_tarjeta_metrica("Papelería Env.", tec_papeleria, "#2d3436"),
        ])
    ], scroll=ft.ScrollMode.AUTO)

    # ==========================================
    # PESTAÑA ESPECIALIZADA (ESP)
    # ==========================================
    dd_periodo_esp = ft.Dropdown(
        label="Selecciona el Periodo",
        options=[ft.dropdown.Option(p) for p in periodos_esp],
        value=periodos_esp[0] if periodos_esp else None,
        width=200
    )

    esp_tot_alumnos = ft.Text("0", size=28, weight=ft.FontWeight.BOLD, color="#a29bfe")
    esp_form_contestado = ft.Text("0", size=28, weight=ft.FontWeight.BOLD, color="#a29bfe")
    esp_pendientes = ft.Text("0", size=28, weight=ft.FontWeight.BOLD, color="#ff4757")

    def actualizar_dashboard_esp(e=None):
        periodo_sel = dd_periodo_esp.value
        total = 0; contestaron = 0; pendientes = 0
        
        for fila in datos_esp[1:]:
            if len(fila) > 0 and str(fila[0]).startswith(str(periodo_sel)):
                total += 1
                if len(fila) > 7 and fila[7].strip().upper() == "SI":
                    contestaron += 1
                if len(fila) > 10 and fila[10].strip() != "Si":
                    pendientes += 1

        esp_tot_alumnos.value = str(total)
        esp_form_contestado.value = str(contestaron)
        esp_pendientes.value = str(pendientes)
        page.update()

    dd_periodo_esp.on_change = actualizar_dashboard_esp

    vista_esp = ft.Column([
        ft.Row([dd_periodo_esp], alignment=ft.MainAxisAlignment.CENTER),
        ft.Row([
            crear_tarjeta_metrica("Pendientes", esp_pendientes, "#1e272e"),
            crear_tarjeta_metrica("Total Alumnos", esp_tot_alumnos, "#2d3436"),
        ]),
        ft.Row([
            crear_tarjeta_metrica("Form. Contestados", esp_form_contestado, "#2d3436"),
        ])
    ], scroll=ft.ScrollMode.AUTO)

    # ==========================================
    # PESTAÑA BUSCADOR
    # ==========================================
    txt_buscar = ft.TextField(label="Buscar por Matrícula o Nombre", expand=True, on_submit=lambda e: buscar_alumno())
    btn_buscar = ft.IconButton(icon=ft.Icons.SEARCH, on_click=lambda e: buscar_alumno())
    lista_resultados = ft.ListView(expand=True, spacing=10)

    def buscar_alumno():
        lista_resultados.controls.clear()
        query = txt_buscar.value.strip().lower()
        if not query:
            page.update()
            return
            
        resultados_encontrados = []
        
        # Buscar en TEC
        for fila in datos_tec[1:]:
            if len(fila) > 2:
                matricula = str(fila[1]).lower()
                nombre = str(fila[2]).lower()
                if query in matricula or query in nombre:
                    carrera = fila[3] if len(fila) > 3 else "Desconocida"
                    estado = fila[8] if len(fila) > 8 else "Pendiente"
                    resultados_encontrados.append(("TEC", fila[1], fila[2], carrera, estado))
                    
        # Buscar en ESP
        for fila in datos_esp[1:]:
            if len(fila) > 2:
                matricula = str(fila[1]).lower()
                nombre = str(fila[2]).lower()
                if query in matricula or query in nombre:
                    oferta = fila[4] if len(fila) > 4 else "Desconocida"
                    estado = fila[10] if len(fila) > 10 else "Pendiente"
                    resultados_encontrados.append(("ESP", fila[1], fila[2], oferta, estado))

        if not resultados_encontrados:
            lista_resultados.controls.append(ft.Text("No se encontraron resultados.", color=ft.Colors.RED_400))
        else:
            for tipo, mat, nom, prog, est in resultados_encontrados:
                color_badge = ft.Colors.BLUE_400 if tipo == "TEC" else ft.Colors.PURPLE_400
                lista_resultados.controls.append(
                    ft.Card(
                        content=ft.Container(
                            content=ft.Column([
                                ft.Row([
                                    ft.Badge(text=tipo, color="white", bgcolor=color_badge),
                                    ft.Text(f"Matrícula: {mat}", weight="bold")
                                ]),
                                ft.Text(f"{nom}", size=16),
                                ft.Text(f"Programa: {prog}", size=12, color=ft.Colors.WHITE_70),
                                ft.Text(f"Estado: {est}", size=12, color=ft.Colors.GREEN_400 if est=="Si" else ft.Colors.AMBER_400),
                            ]),
                            padding=15
                        )
                    )
                )
        page.update()

    vista_buscador = ft.Column([
        ft.Row([txt_buscar, btn_buscar]),
        lista_resultados
    ], expand=True)

    # ==========================================
    # ENSAMBLAJE DE PESTAÑAS (Flet 0.25+)
    # ==========================================
    t = ft.Tabs(
        selected_index=0,
        length=3,
        expand=True,
        content=ft.Column(
            expand=True,
            controls=[
                ft.TabBar(
                    tabs=[
                        ft.Tab(label="TÉCNICA", icon=ft.Icons.SCHOOL),
                        ft.Tab(label="ESPECIALIZADA", icon=ft.Icons.WORKSPACE_PREMIUM),
                        ft.Tab(label="BUSCADOR", icon=ft.Icons.PERSON_SEARCH),
                    ]
                ),
                ft.TabBarView(
                    expand=True,
                    controls=[
                        ft.Container(content=vista_tec, padding=10),
                        ft.Container(content=vista_esp, padding=10),
                        ft.Container(content=vista_buscador, padding=10),
                    ]
                )
            ]
        )
    )

    page.add(t)
    
    # Inicializar datos al abrir
    actualizar_dashboard_tec()
    actualizar_dashboard_esp()

ft.app(target=main)