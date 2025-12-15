# CutPlanner
Gestión de inventario y planificación de cortes para ventanas, persianas, mosquiteras y similares.
# 🏭 CutPlanner - Sistema de Optimización de Cortes e Inventario

**CutPlanner** es una solución avanzada de gestión de producción para talleres de carpintería de aluminio y PVC. Desarrollado en **Google Apps Script**, se integra directamente en Google Sheets para calcular la forma más eficiente de cortar perfiles, gestionar el stock real y reutilizar retales (sobrantes útiles).

## 🚀 Características Principales

* **Motor de Optimización Inteligente:** Utiliza un algoritmo de búsqueda en haz (*Beam Search*) con "Paciencia Dinámica" para encontrar el equilibrio entre el desperdicio mínimo y el tiempo de cálculo.
* **Lógica de "Coste Progresivo":** El sistema prioriza gastar retales viejos antes que abrir barras nuevas.
* **Gestión de Retales:** Identifica automáticamente los sobrantes útiles, les asigna un ID único y los guarda en el inventario para futuros trabajos.
* **Filtrado Avanzado de Perfiles:** Soporte nativo para lógicas complejas como perfiles **H10 vs H18** y diferenciación de marcos (Izquierda/Derecha en Compactos).
* **Dos Modos de Operación:**
    * `🧮 Modo Solo`: Simulación de cortes para presupuestos sin afectar el stock.
    * `▶️ Modo Producción`: Ejecuta cortes, descuenta material del inventario y guarda retales.

## 📂 Estructura del Proyecto

El código está modularizado en los siguientes archivos:

* **`Main.gs`**: Punto de entrada. Crea el menú "Gestión de Inventario" y orquesta los modos de ejecución.
* **`Optimizer.gs`**: El "cerebro" del sistema. Contiene la lógica heurística, el sistema de puntuación (Juez) y el control de *timeout* de Google.
* **`Reader.gs`**: Se encarga de leer la hoja de datos, filtrar por tipos de ventana (Compacto, 2 Hojas) y parsear la configuración.
* **`InventoryManager.gs`**: Gestiona las altas/bajas en la hoja 'Inventario', limpieza de filas vacías y generación de IDs para retales.
* **`Config.gs`**: Define los perfiles (colores, columnas, multiplicadores) y parámetros globales del sistema.
* **`Workflow.gs`**: Controla el flujo de la hoja de cálculo (marcar casillas de verificación como "Hecho").
* **`Visuals.gs` / `ReportWriter.gs`**: (Opcionales) Generación de informes visuales y resumen de materiales.

## ⚙️ Configuración

El sistema se alimenta de la hoja **`Configuración`**, donde se definen parámetros críticos:

| Parámetro | Descripción |
| :--- | :--- |
| `Longitud_Barra_Nueva` | Largo estándar del material (ej. 6500mm). |
| `Saneo` | Material descartado al inicio de la barra (ej. 40mm). |
| `Desperdicio_por_Corte` | Grosor del disco de corte (ej. 4mm). |
| `Limite_Retal` | Longitud mínima para guardar un sobrante (ej. 1000mm). |
| `Limite_Desechable` | Longitud máxima para considerar basura (ej. 200mm). |

### Definición de Perfiles (`Config.gs`)
Los perfiles se configuran en el objeto `TODOS_PERFILES`. Cada perfil incluye:
* `id`: Identificador único para el inventario.
* `columna`: Columna de la hoja de datos de donde lee la medida.
* `tipoRequerido`: (Opcional) Filtra cortes según el tipo de ventana (ej. `["Compacto"]` o `["2 Hojas"]`).

## 🛠️ Instalación y Uso

1.  **Requisitos:** Una hoja de cálculo de Google con las pestañas: `Inventario`, `Configuración`, `Corredera_GP-60` (Datos) y `Plan_Cortes` (Histórico).
2.  **Instalación:** Copiar los archivos `.gs` al editor de Apps Script asociado a la hoja.
3.  **Ejecución:**
    * Recargar la hoja de cálculo.
    * Aparecerá un nuevo menú en la barra superior: **🏭 Gestión de Inventario**.

### Opciones del Menú

* **▶️ PRODUCCIÓN (Usa Stock y Actualiza):**
    * Lee los cortes pendientes (checkbox desmarcado).
    * Busca material en el inventario real. Si falta, sugiere compra.
    * Genera el plan de corte, actualiza el stock y marca los trabajos como realizados.
* **🧮 MODO SOLO (Solo Calcular):**
    * Calcula las necesidades teóricas usando stock infinito.
    * Ideal para estimar materiales antes de aceptar un pedido.
* **🔄 Recalcular Todo / 👁️ Cambiar Vista:** Herramientas de utilidad para refrescar datos.

## 🧠 Lógica del Algoritmo

El optimizador (`Optimizer.gs`) puntúa cada posible corte basándose en:
1.  **Aprovechamiento:** ¿Cuánto material sobra? (Prefiere sobrantes muy grandes o nulos).
2.  **Penalización de "Zona Muerta":** Evita generar retales inútiles (entre 20cm y 1m).
3.  **Prioridad de Retales:** Penaliza el uso de barras nuevas si existen retales compatibles en el inventario.

---
*CutPlanner v2.0 - Optimización de procesos para carpintería.*
