# 🏥 SaludSV Pro
### Tu Organizador Personal de Salud — El Salvador

![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1+-green)
![Excel](https://img.shields.io/badge/Excel-2016+-darkgreen?logo=microsoft-excel)
![License](https://img.shields.io/badge/License-MIT-yellow)
![Claude AI](https://img.shields.io/badge/Built%20with-Claude%20AI-orange)
![El Salvador](https://img.shields.io/badge/País-El%20Salvador-blue)

> **"Tu salud, organizada. Tu vida, protegida."**

SaludSV Pro es una herramienta de salud personal generada con Python que produce un archivo Excel interactivo y profesional, diseñado específicamente para la población salvadoreña. Permite gestionar citas médicas, medicamentos, rutinas de ejercicio, dieta e historial clínico en un solo documento portable que funciona sin conexión a internet.

---

## 📋 Tabla de contenidos

- [El problema que resuelve](#-el-problema-que-resuelve)
- [Demo visual](#-demo-visual)
- [Características principales](#-características-principales)
- [Módulos del sistema](#-módulos-del-sistema)
- [Requisitos](#-requisitos)
- [Instalación y uso](#-instalación-y-uso)
- [Archivos generados](#-archivos-generados)
- [Integración con Google Calendar](#-integración-con-google-calendar)
- [Stack tecnológico](#-stack-tecnológico)
- [Estructura del proyecto](#-estructura-del-proyecto)
- [Créditos](#-créditos)

---

## 🇸🇻 El problema que resuelve

En El Salvador, millones de personas con condiciones crónicas como diabetes e hipertensión:

- Manejan sus medicamentos en papelitos o de memoria
- Olvidan citas médicas por falta de recordatorios
- No tienen su historial clínico organizado para compartir con médicos
- Pierden control de sus rutinas de ejercicio y dieta

**SaludSV Pro centraliza toda esa información** en un solo archivo descargable, sin necesidad de internet, aplicaciones especiales ni suscripciones — solo Excel.

---

## 🎬 Demo visual

```
python salud_sv_pro.py
```

```
============================================================
   SaludSV Pro — Generador de Organizador de Salud
   Versión 1.0 | El Salvador 2026 | Powered by Claude AI
============================================================

✓ Carpetas de output creadas
→ Creando hoja: INICIO...
→ Creando hoja: CITAS MÉDICAS...
→ Creando hoja: MEDICAMENTOS...
→ Creando hoja: EJERCICIOS...
→ Creando hoja: DIETA...
→ Creando hoja: HISTORIAL CLÍNICO...
→ Creando hoja: INSTRUCCIONES...
✓ Generando archivos .ics...
  ✓ cita_001.ics
  ✓ cita_002.ics
  ✓ cita_003.ics
✓ README.md generado

============================================================
  ¡SaludSV Pro generado exitosamente!
  → SaludSV_Pro.xlsx
  → calendario/ (3 archivos .ics)
============================================================
```

---

## ✨ Características principales

| Característica | Descripción |
|---|---|
| 📊 **7 módulos integrados** | Dashboard, Citas, Medicamentos, Ejercicios, Dieta, Historial, Instrucciones |
| 🎨 **Formato condicional** | Colores automáticos según urgencia y estado |
| 📅 **Exportación a calendario** | Archivos .ics con doble recordatorio para Google Calendar y Outlook |
| 💊 **Alertas de refill** | Aviso automático cuando un medicamento está por agotarse |
| 🏋️ **Aviso de cambio de rutina** | Recordatorio automático cada 6 semanas |
| 🥗 **Tabla nutricional salvadoreña** | 20 alimentos típicos de El Salvador con valores nutricionales |
| 📋 **Historial clínico portable** | Documento listo para compartir con cualquier médico |
| ⚡ **Funciona offline** | No requiere internet ni instalaciones adicionales para el usuario |
| 🔒 **Sin macros** | Compatible con cualquier versión de Excel, sin riesgos de seguridad |

---

## 📁 Módulos del sistema

### 🏥 Módulo 1: INICIO (Dashboard)

Panel principal del organizador. Es la primera hoja que se ve al abrir el archivo.

**Funciones:**
- **Datos del paciente:** nombre, fecha de nacimiento, tipo de sangre, teléfono, dirección, seguro médico (ISSS, COSENA, Bienestar Magisterial, Privado, Ninguno)
- **Cálculo automático de edad** a partir de la fecha de nacimiento
- **Panel de alertas activas:** muestra la próxima cita, medicamentos a refill, días para cambio de rutina y fecha actual
- **Registro de alergias** con nivel de severidad (Leve, Moderada, Severa, Anafilaxia)
- **Condiciones crónicas** con estado de control (Controlada, En tratamiento, Monitoreo, Remisión)
- Datos de ejemplo pre-llenados para facilitar el inicio

---

### 📅 Módulo 2: CITAS MÉDICAS

Registro completo de todas las consultas médicas con sistema visual de estados.

**Funciones:**
- Registro de: fecha, hora, médico, especialidad, clínica/hospital, dirección, motivo y costo
- **Dropdown de especialidades** con 22 opciones (Medicina General, Cardiología, Endocrinología, Psiquiatría, y más)
- **Dropdown de estados:** Pendiente, Completada, Cancelada, Reprogramada, En espera de confirmación
- **Código de colores automático por fila:**
  - 🟢 Verde — Cita pendiente con fecha futura
  - 🟡 Amarillo brillante — La cita es hoy
  - 🟠 Naranja — La cita es mañana o pasado mañana
  - ⬛ Gris — Cita completada
  - 🔴 Rojo — Cita cancelada
- **Total de gastos** calculado automáticamente
- Autofilter y encabezados congelados para navegación fácil
- 3 citas de ejemplo pre-cargadas con fechas relativas al día actual

---

### 💊 Módulo 3: MEDICAMENTOS

Control completo del régimen farmacológico con sistema de alertas de reabastecimiento.

**Funciones:**
- Registro de: nombre, presentación, dosis, vía de administración, frecuencia, horario, cantidad actual
- **Dropdowns completos:**
  - Presentación: Tableta, Cápsula, Jarabe, Inyectable, Crema, Parche, Gotas, Inhalador y más
  - Vía: Oral, Tópica, Intramuscular, Intravenosa, Subcutánea, Inhalatoria y más
  - Frecuencia: Cada 4h, 6h, 8h, 12h, 24h, Semanal, Mensual, Según necesidad
  - Tomar con: Con alimentos, En ayunas, Con leche, Sin restricción y más
- **Cálculo automático de días restantes** según la frecuencia de cada medicamento
- **Sistema de alertas por color basado en días restantes:**
  - 🔴 Rojo urgente — 3 días o menos (Refill inmediato)
  - 🟠 Naranja — 7 días o menos
  - 🟡 Amarillo — 14 días o menos
  - 🟢 Verde — Stock suficiente (más de 14 días)
- Indicador de medicamentos crónicos vs temporales
- Panel lateral de **horario del día** (mañana, tarde, noche)
- 3 medicamentos de ejemplo: Metformina, Losartán y Vitamina D3

---

### 💪 Módulo 4: EJERCICIOS

Planificador de rutinas de entrenamiento con contador automático de cambio de rutina.

**Funciones:**
- **Panel de rutina activa:**
  - Nombre de la rutina
  - Fecha de inicio
  - Semana actual de rutina (calculada automáticamente)
  - Fecha de cambio de rutina (fecha inicio + 42 días)
  - Contador de días restantes para el cambio
  - **Alerta visual dinámica:** roja si quedan ≤7 días, naranja si ≤14 días, verde si >14 días
  - Datos de contacto del entrenador personal
- **Tabla de ejercicios** con: grupo muscular, nombre del ejercicio, series, repeticiones/tiempo, peso, descanso, días de la semana y notas técnicas
- **Dropdown de grupos musculares:** Pecho, Espalda, Hombros, Bíceps, Tríceps, Abdomen/Core, Glúteos, Cuádriceps, Isquiotibiales, Pantorrillas, Cardio, Movilidad y más
- **Registro mensual** de adherencia con tabla 4 semanas × 7 días
- Rutina full body de 3 días pre-cargada como ejemplo (10 ejercicios)

---

### 🥗 Módulo 5: DIETA Y NUTRICIÓN

Planificador semanal de alimentación con tabla nutricional contextualizada para El Salvador.

**Funciones:**
- **Panel de metas nutricionales diarias:** calorías, proteína, carbohidratos, grasas y vasos de agua
- **Plan semanal de comidas** (Lunes a Domingo) con 5 tiempos de comida:
  - Desayuno, Snack mañana, Almuerzo, Snack tarde, Cena
- Registro de consumo de agua diario (dropdown 1-10 vasos)
- Notas diarias de observaciones
- **Tabla nutricional de 20 alimentos típicos salvadoreños** con valores por porción:
  - Pupusa de queso, Pupusa revuelta, Arroz blanco, Frijoles rojos
  - Pollo a la plancha, Carne de res asada, Huevo, Pan francés
  - Plátano fresco, Plátano frito, Yuca cocida, Tortilla de maíz
  - Aguacate, Leche entera, Queso duro seco, Atún en agua
  - Tomate, Zanahoria, Naranja, Yogur natural
- Ejemplo completo del lunes pre-llenado como referencia

---

### 📋 Módulo 6: HISTORIAL CLÍNICO

Expediente médico personal portable, listo para compartir con cualquier médico.

**Funciones:**
- **Datos médicos básicos:** grupo sanguíneo, factor Rh, peso, talla
- **Cálculo automático de IMC** con clasificación OMS:
  - 🔵 Azul — Bajo peso (< 18.5)
  - 🟢 Verde — Normal (18.5 - 24.9)
  - 🟡 Amarillo — Sobrepeso (25 - 29.9)
  - 🔴 Rojo — Obesidad (≥ 30)
- **Seguimiento histórico de signos vitales:** peso, presión arterial sistólica/diastólica, glucosa en ayunas y frecuencia cardíaca con fecha
- **Registro de vacunas** con fecha de aplicación y próxima dosis (incluye vacunas comunes en El Salvador: COVID-19, Influenza, Tétanos, Hepatitis B)
- **Historial de cirugías** y procedimientos quirúrgicos
- **Antecedentes familiares** (Padre, Madre, Abuelos paternos y maternos, Hermanos)
- **Directorio de médicos tratantes** con especialidad, clínica, teléfono y frecuencia de visita

---

### 📖 Módulo 7: INSTRUCCIONES DE USO

Manual visual integrado en el archivo para que cualquier usuario pueda usarlo sin asistencia.

**Funciones:**
- Guía de uso para cada uno de los 6 módulos
- Tabla visual de **guía de colores** con muestras reales de cada alerta
- **Pasos detallados** para importar citas a Google Calendar y Outlook
- Recomendaciones de uso y buenas prácticas
- Créditos y versión del aplicativo

---

## 💻 Requisitos

- Python 3.8 o superior
- Microsoft Excel 2016+ o LibreOffice Calc 7+
- Windows, macOS o Linux

---

## 🚀 Instalación y uso

### 1. Clonar el repositorio
```bash
git clone https://github.com/tu-usuario/saludsv-pro.git
cd saludsv-pro
```

### 2. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 3. Generar el organizador
```bash
python salud_sv_pro.py
```

### 4. Abrir el archivo
El archivo `SaludSV_Pro.xlsx` se genera en la carpeta `output/`. Ábrelo con Excel o LibreOffice Calc.

---

## 📂 Archivos generados

```
output/
├── SaludSV_Pro.xlsx          # Organizador principal (7 hojas)
├── README.md                 # Instrucciones de uso
└── calendario/
    ├── cita_001.ics          # Cita 1 para Google Calendar
    ├── cita_002.ics          # Cita 2 para Google Calendar
    └── cita_003.ics          # Cita 3 para Google Calendar
```

---

## 📅 Integración con Google Calendar

Cada cita médica registrada genera automáticamente un archivo `.ics` con **dos recordatorios automáticos:**

- ⏰ **1 día antes** — "Recordatorio: Mañana tienes cita con [Doctor]"
- ⏰ **2 horas antes** — "Tu cita con [Doctor] es en 2 horas - [Hospital]"

### Cómo importar una cita:
1. Ve a la carpeta `output/calendario/`
2. Haz doble clic en `cita_001.ics`
3. Google Calendar o Outlook se abre automáticamente
4. Confirma la importación
5. Los recordatorios quedan configurados automáticamente

Compatible con: **Google Calendar**, **Microsoft Outlook**, **Apple Calendar** y cualquier aplicación que soporte el estándar iCalendar (RFC 5545).

---

## 🛠️ Stack tecnológico

| Tecnología | Versión | Uso |
|---|---|---|
| **Python** | 3.8+ | Lenguaje principal |
| **openpyxl** | 3.1+ | Generación y formato del Excel |
| **icalendar** | 5.0+ | Generación de archivos .ics |
| **pytz** | 2023.3+ | Zona horaria America/El_Salvador |
| **python-dateutil** | 2.8+ | Manejo avanzado de fechas |
| **Claude AI** | Sonnet 4 | Generación del código vía Claude Code |

---

## 🗂️ Estructura del proyecto

```
saludsv-pro/
├── salud_sv_pro.py          # Script principal generador
├── CLAUDE.md                # Especificación técnica completa
├── requirements.txt         # Dependencias Python
├── README.md                # Este archivo
└── output/                  # Archivos generados (auto-creada)
    ├── SaludSV_Pro.xlsx
    └── calendario/
```

---

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Para cambios importantes:

1. Haz fork del repositorio
2. Crea una rama: `git checkout -b feature/nueva-funcionalidad`
3. Commit: `git commit -m 'Agrega nueva funcionalidad'`
4. Push: `git push origin feature/nueva-funcionalidad`
5. Abre un Pull Request

---

## 📜 Licencia

Este proyecto está bajo la licencia MIT. Consulta el archivo `LICENSE` para más detalles.

---

## 👨‍💻 Créditos

Desarrollado como proyecto final del **Bootcamp de Claude — Anthropic**  
El Salvador, 2026

**Herramientas utilizadas:**
- Desarrollado con [Claude Code](https://claude.ai/code) — Anthropic
- Especificado y construido con [Claude Sonnet](https://anthropic.com)

---

<div align="center">
  
**SaludSV Pro v1.0**  
*Hecho con ❤️ para El Salvador*

</div>
