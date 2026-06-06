# CLAUDE.md — SaludSV Pro
## Organizador Personal de Salud para El Salvador
### Versión 1.0 | Desarrollado con Claude AI | 2026

---

## DESCRIPCIÓN DEL PROYECTO

SaludSV Pro es una aplicación Python que genera un archivo Excel interactivo y
profesional para la gestión integral de salud personal. Diseñado específicamente
para la población salvadoreña, permite organizar citas médicas, medicamentos,
rutinas de ejercicio, dieta e historial clínico en un solo documento portable,
descargable y accesible sin conexión a internet.

**Problema que resuelve:** En El Salvador, muchos pacientes manejan múltiples
médicos, medicamentos crónicos (diabetes, hipertensión) y no tienen un sistema
centralizado para su información de salud. SaludSV Pro es ese sistema.

---

## ARCHIVOS A CREAR

```
saludsv_pro/
├── CLAUDE.md                    ← Este archivo
├── requirements.txt             ← Dependencias Python
├── salud_sv_pro.py              ← Script principal (GENERAR ESTE)
├── README.md                    ← Instrucciones de uso (GENERAR ESTE)
└── output/
    ├── SaludSV_Pro.xlsx         ← Artefacto Excel final
    └── calendario/              ← Archivos .ics por cita
        ├── cita_001.ics
        ├── cita_002.ics
        └── cita_003.ics
```

---

## DEPENDENCIAS (requirements.txt)

```
openpyxl>=3.1.0
icalendar>=5.0.0
pytz>=2023.3
python-dateutil>=2.8.2
```

---

## PALETA DE COLORES OFICIAL

Usar EXACTAMENTE estos códigos hexadecimales en todo el archivo:

| Nombre           | Hex       | Uso                                      |
|------------------|-----------|------------------------------------------|
| COLOR_PRIMARIO   | #1B6B3A   | Headers principales, pestaña INICIO      |
| COLOR_SECUNDARIO | #0F4C8A   | Sub-headers, pestaña CITAS               |
| COLOR_ACENTO     | #27AE60   | Elementos OK/verde, pestaña DIETA        |
| ALERTA_ROJO      | #C0392B   | Refill urgente, cancelado, pestaña MEDS  |
| ALERTA_NARANJA   | #E67E22   | Advertencia pronto (≤7 días)             |
| ALERTA_AMARILLO  | #F1C40F   | Advertencia moderada (≤14 días)          |
| ALERTA_VERDE     | #2ECC71   | Todo en orden (>14 días)                 |
| FONDO_CLARO      | #F8F9FA   | Filas alternas, fondo secciones          |
| FONDO_GRIS       | #ECF0F1   | Áreas de solo lectura                    |
| TEXTO_BLANCO     | #FFFFFF   | Texto sobre fondos oscuros               |
| TEXTO_OSCURO     | #2C3E50   | Texto principal sobre fondo claro        |
| MORADO           | #8E44AD   | Pestaña EJERCICIOS                       |
| GRIS_OSCURO      | #2C3E50   | Pestaña HISTORIAL                        |
| GRIS_MEDIO       | #7F8C8D   | Pestaña INSTRUCCIONES                    |

---

## TIPOGRAFÍA

- Título principal de hoja: Calibri, 16pt, Bold, color TEXTO_BLANCO
- Título de sección: Calibri, 13pt, Bold, color TEXTO_BLANCO
- Encabezado de columna: Calibri, 11pt, Bold, color TEXTO_BLANCO
- Datos de celda: Calibri, 11pt, Normal, color TEXTO_OSCURO
- Notas/instrucciones: Calibri, 10pt, Italic, color #7F8C8D

---

## HOJA 1: INICIO (Dashboard)

### Configuración general
- Nombre de pestaña: "🏥 INICIO"
- Color de pestaña: #1B6B3A
- Sin líneas de cuadrícula (sheet.sheet_view.showGridLines = False)
- Zoom: 90%
- Congelar paneles: fila 4 (ws.freeze_panes = "A4")

### Bloque A — Banner principal (filas 1-3, columnas A:H)
- Merge A1:H3
- Fondo: #1B6B3A
- Texto centrado (horizontal y vertical): "SaludSV Pro"
- Fuente: Calibri 22pt, Bold, Blanco
- Segunda línea (misma celda con salto): "Tu Organizador Personal de Salud"
- Fuente segunda línea: Calibri 12pt, Normal, #A9DFBF

### Bloque B — Datos del paciente (filas 5-15, columnas A:D)
Título "DATOS DEL PACIENTE" en A5:D5, fondo #0F4C8A, texto blanco, bold 12pt.

Campos (etiqueta en col A bold, dato en B:D merge):
| Fila | Etiqueta (col A)        | Dato (col B-D)                              |
|------|------------------------|---------------------------------------------|
| 6    | Nombre completo:       | [campo editable vacío]                      |
| 7    | Fecha de nacimiento:   | [campo fecha DD/MM/AAAA]                    |
| 8    | Edad:                  | Fórmula: =SI(B7="","",SIFECHA(B7,HOY(),"Y")&" años") |
| 9    | Tipo de sangre:        | Dropdown: O+,O-,A+,A-,B+,B-,AB+,AB-        |
| 10   | Teléfono personal:     | [campo editable]                            |
| 11   | Dirección:             | [campo editable]                            |
| 12   | Municipio/Ciudad:      | [campo editable, sugerencia: San Salvador]  |
| 13   | Contacto emergencia:   | [campo editable]                            |
| 14   | Tel. emergencia:       | [campo editable]                            |
| 15   | Seguro médico:         | Dropdown: ISSS,COSENA,BIENESTAR MAGISTERIAL,PRIVADO,NINGUNO |

### Bloque C — Alertas activas (filas 5-15, columnas F:H)
Título "ALERTAS ACTIVAS" en F5:H5, fondo #C0392B, texto blanco, bold 12pt.

Celdas de alerta (cada una con fondo suave y texto informativo):
| Fila | Contenido                                              |
|------|--------------------------------------------------------|
| 6    | "Próxima cita:" (label bold) + referencia a hoja CITAS |
| 8    | "Medicamentos a refill:" + conteo de meds críticos     |
| 10   | "Cambio de rutina en:" + días restantes de EJERCICIOS  |
| 12   | "Hoy:" + =TEXTO(HOY(),"DD/MM/AAAA")                   |
| 14   | "Versión: SaludSV Pro 1.0"                            |

Usar fondo #FEF9E7 y borde #F39C12 para las celdas de alerta.

### Bloque D — Alergias (filas 17-22, columnas A:E)
Título "ALERGIAS CONOCIDAS" en A17:E17, fondo #E74C3C, texto blanco.
Encabezados en fila 18: ALERGIA | SEVERIDAD | REACCIÓN | FECHA DIAGNÓSTICO | NOTAS
Dropdown SEVERIDAD: Leve, Moderada, Severa, Anafilaxia
Datos ejemplo fila 19: Polen, Leve, Estornudos, 01/01/2020, Antihistamínico si es necesario

### Bloque E — Condiciones crónicas (filas 24-29, columnas A:E)
Título "CONDICIONES CRÓNICAS" en A24:E24, fondo #8E44AD, texto blanco.
Encabezados fila 25: CONDICIÓN | FECHA DIAGNÓSTICO | MÉDICO TRATANTE | ESTADO | MEDICACIÓN
Dropdown ESTADO: Controlada, En tratamiento, Monitoreo, Remisión
Datos ejemplo fila 26: Diabetes tipo 2, 15/03/2018, Dr. García, Controlada, Metformina

### Widths de columnas INICIO
A:22, B:16, C:16, D:16, E:2(separador), F:20, G:16, H:16

---

## HOJA 2: CITAS MÉDICAS

### Configuración
- Nombre pestaña: "📅 CITAS MÉDICAS"
- Color pestaña: #0F4C8A
- Sin gridlines
- Autofilter activado en headers
- Congelar fila de headers
- Orientación impresión: Landscape

### Banner (filas 1-2, A:K)
Merge A1:K2, fondo #0F4C8A
Texto: "CITAS MÉDICAS | Registro y seguimiento de consultas"
Sub: "Verde=Pendiente | Naranja=Mañana/Pasado mañana | Gris=Completada | Rojo=Cancelada"

### Tabla de citas (desde fila 4)
Header en fila 4, fondo #0F4C8A, texto blanco, bold 11pt, centrado, alto 30px.

| Col | Header              | Ancho | Formato           |
|-----|---------------------|-------|-------------------|
| A   | #                   | 5     | Número entero     |
| B   | FECHA               | 13    | DD/MM/AAAA        |
| C   | HORA                | 8     | HH:MM             |
| D   | DOCTOR/A            | 24    | Texto             |
| E   | ESPECIALIDAD        | 22    | Dropdown          |
| F   | CLÍNICA / HOSPITAL  | 28    | Texto             |
| G   | DIRECCIÓN           | 30    | Texto             |
| H   | MOTIVO              | 30    | Texto             |
| I   | ESTADO              | 14    | Dropdown          |
| J   | COSTO ($)           | 11    | Número 2 decimales|
| K   | NOTAS               | 35    | Texto             |

Dropdown ESPECIALIDAD:
Medicina General,Cardiología,Dermatología,Endocrinología,Gastroenterología,
Ginecología,Medicina Interna,Nefrología,Neumología,Neurología,Odontología,
Oftalmología,Oncología,Ortopedia,Otorrinolaringología,Pediatría,Psiquiatría,
Reumatología,Traumatología,Urología,Nutrición,Fisioterapia

Dropdown ESTADO:
Pendiente,Completada,Cancelada,Reprogramada,En espera de confirmación

### Formato condicional (aplicar a filas 5:104)
- Toda la fila ROJO (#FADBD8, texto #922B21): si col I = "Cancelada"
- Toda la fila GRIS (#D5DBDB, texto #717D7E): si col I = "Completada"
- Toda la fila NARANJA (#FDEBD0, texto #935116): si col B = HOY()+1 O HOY()+2
- Toda la fila VERDE (#D5F5E3, texto #1E8449): si col I = "Pendiente" Y col B >= HOY()+3
- Toda la fila AMARILLO BRILLANTE (#FEF9E7, texto #9A7D0A): si col B = HOY()
Filas alternas sin formato condicional: fondo #F8F9FA

### 3 citas de ejemplo (filas 5-7)
Usar fechas relativas calculadas con datetime.today() + timedelta:
1. HOY+7, 09:00, Dr. Carlos Mendoza, Medicina General, Hospital Nacional Rosales, 
   Final Blvd. Los Héroes, San Salvador, Control mensual de presión arterial, Pendiente, 25.00
2. HOY+14, 14:30, Dra. María García, Endocrinología, Clínica Médica San Salvador,
   75 Av. Norte, Col. Escalón, San Salvador, Control diabetes - HbA1c, Pendiente, 45.00
3. HOY+30, 10:00, Dr. José Rivera, Cardiología, ISSS Hospital Médico Quirúrgico,
   Blvd. Constitución, San Salvador, Ecocardiograma de control, Pendiente, 0.00

### Fila de totales
Después de los datos, una fila con: "TOTAL GASTADO:" + =SUMA(J5:J104) + formato moneda

---

## HOJA 3: MEDICAMENTOS

### Configuración
- Nombre pestaña: "💊 MEDICAMENTOS"
- Color pestaña: #C0392B
- Sin gridlines, landscape printing

### Banner (filas 1-2)
Fondo #C0392B, texto: "MEDICAMENTOS Y RECORDATORIOS"
Sub-línea: "Rojo=Refill urgente (≤3 días) | Naranja=Refill pronto (≤7 días) | Amarillo=Atención (≤14 días) | Verde=Stock OK"

### Tabla principal (desde fila 4)
Header fila 4, fondo #C0392B, texto blanco, bold.

| Col | Header                  | Ancho | Tipo/Formato                        |
|-----|-------------------------|-------|-------------------------------------|
| A   | #                       | 5     | Número                              |
| B   | MEDICAMENTO             | 26    | Texto                               |
| C   | PRESENTACIÓN            | 18    | Dropdown                            |
| D   | DOSIS                   | 15    | Texto (ej: "500mg", "1 tableta")    |
| E   | VÍA ADMINISTRACIÓN      | 16    | Dropdown                            |
| F   | FRECUENCIA              | 20    | Dropdown                            |
| G   | HORARIO                 | 22    | Texto (ej: "7am - 1pm - 7pm")       |
| H   | TOMAR CON               | 18    | Dropdown                            |
| I   | CANTIDAD ACTUAL         | 16    | Número entero ≥ 0                   |
| J   | ALERTA MÍNIMA           | 16    | Número entero ≥ 0                   |
| K   | DÍAS RESTANTES          | 15    | Fórmula (ver abajo)                 |
| L   | ES CRÓNICO              | 12    | Dropdown: Sí, No                    |
| M   | MÉDICO PRESCRIPTOR      | 24    | Texto                               |
| N   | FECHA INICIO            | 13    | Fecha                               |
| O   | FECHA FIN               | 15    | Texto (fecha o "Indefinido")        |
| P   | OBSERVACIONES           | 32    | Texto                               |

Dropdown PRESENTACIÓN:
Tableta,Cápsula,Jarabe,Suspensión,Inyectable,Crema/Ungüento,Parche,Gotas,Inhalador,Supositorio,Polvo,Ampolla

Dropdown VÍA:
Oral,Tópica,Intramuscular,Intravenosa,Subcutánea,Inhalatoria,Oftálmica,Ótica,Sublingual,Rectal,Nasal

Dropdown FRECUENCIA:
Cada 4 horas,Cada 6 horas,Cada 8 horas,Cada 12 horas,Cada 24 horas,
Dos veces por semana,Una vez por semana,Una vez al mes,Según necesidad

Dropdown TOMAR CON:
Con agua,Con alimentos,En ayunas,Con leche,Sin restricción,Alejado de comidas,Con jugo de naranja

### Fórmula DÍAS RESTANTES (columna K)
Para cada fila n (comenzando en 5):
=SI(I{n}=0,"Sin stock",SI(F{n}="Cada 4 horas",REDONDEAR(I{n}/6,0),SI(F{n}="Cada 6 horas",REDONDEAR(I{n}/4,0),SI(F{n}="Cada 8 horas",REDONDEAR(I{n}/3,0),SI(F{n}="Cada 12 horas",REDONDEAR(I{n}/2,0),SI(F{n}="Cada 24 horas",I{n},SI(F{n}="Dos veces por semana",I{n}/2*7,SI(F{n}="Una vez por semana",I{n}*7,I{n}))))))))

### Formato condicional (basado en col K, filas 5:54)
- ROJO URGENTE (#FADBD8, texto #922B21, bold): col K <= 3 Y col K <> "Sin stock"
- NARANJA (#FDEBD0, texto #935116): col K > 3 Y col K <= 7
- AMARILLO (#FEF9E7, texto #9A7D0A): col K > 7 Y col K <= 14
- VERDE (#D5F5E3, texto #1E8449): col K > 14
- NEGRO sobre rojo intenso (#C0392B fondo, blanco texto): col K = "Sin stock" o col K = 0

### 3 medicamentos de ejemplo (filas 5-7)
1. Metformina 500mg, Tableta, 500mg, Oral, Cada 12 horas, 7am-7pm, Con alimentos,
   45 tabletas, alerta 10, [fórmula días], Sí, Dr. María García, 15/01/2023, Indefinido,
   Control glucemia en diabetes tipo 2. No suspender sin consultar médico.
2. Losartán 50mg, Tableta, 50mg, Oral, Cada 24 horas, 8am, Con agua,
   20 tabletas, alerta 7, [fórmula], Sí, Dr. Carlos Mendoza, 20/03/2022, Indefinido,
   Para hipertensión arterial. Tomar a la misma hora cada día.
3. Vitamina D3 1000UI, Cápsula, 1000UI, Oral, Cada 24 horas, 8am, Con alimentos,
   60 cápsulas, alerta 10, [fórmula], No, Dr. José Rivera, 01/06/2026, 01/12/2026,
   Suplemento por deficiencia diagnosticada.

### Panel horario del día (columnas R:U, filas 4-20)
Título "HORARIO DEL DÍA" en R4:U4, fondo #8E44AD, texto blanco.
Secciones visuales:
- R5:U8  "MAÑANA (6am-12pm)" fondo #FEF9E7 + instrucción manual de llenado
- R9:U12 "TARDE (12pm-6pm)" fondo #EBF5FB
- R13:U16 "NOCHE (6pm-12am)" fondo #F0F3F4
- R17:U20 "NOTA:" fondo #FDFEFE, texto: "Complete según sus medicamentos activos"

---

## HOJA 4: RUTINA DE EJERCICIOS

### Configuración
- Nombre pestaña: "💪 EJERCICIOS"
- Color pestaña: #8E44AD
- Sin gridlines

### Banner (filas 1-2)
Fondo #8E44AD, texto: "RUTINA DE EJERCICIOS"
Sub: "Mantén tu rutina activa — recuerda cambiarla cada 6 semanas para mejores resultados"

### Panel de info de rutina (filas 4-10, columnas A:F)
Título "MI RUTINA ACTUAL" en A4:F4, fondo #6C3483, texto blanco.

Campos:
| Fila | Etiqueta                     | Valor/Fórmula                                   |
|------|-----------------------------|-------------------------------------------------|
| 5    | Nombre de la rutina:         | [campo editable — ej: "Full Body Intermedio"]   |
| 6    | Fecha inicio de rutina:      | [campo fecha — ej: fecha de hoy - 14 días]      |
| 7    | Semana actual de rutina:     | =SI(B6="","",ENTERO((HOY()-B6)/7)+1)&" de 6"   |
| 8    | Fecha de cambio de rutina:   | =SI(B6="","",B6+42)  (formato fecha)            |
| 9    | Días para cambio:            | =SI(B6="","",MAX(0,B6+42-HOY()))&" días"        |
| 10   | Entrenador personal:         | [campo editable]                                |

Alerta visual en H9:K9:
- Si días para cambio <= 7: fondo #C0392B, texto blanco, bold: "⚠ ¡Es hora de renovar tu rutina!"
- Si días para cambio <= 14: fondo #E67E22, texto blanco: "Prepara tu nueva rutina pronto"
- Si días para cambio > 14: fondo #27AE60, texto blanco: "Rutina en curso ✓"
(Implementar con 3 celdas con condicional manual o formato condicional en esa celda)

### Tabla de ejercicios (desde fila 12)
Header fila 12, fondo #8E44AD, texto blanco, bold.

| Col | Header            | Ancho | Tipo                                      |
|-----|-------------------|-------|-------------------------------------------|
| A   | #                 | 5     | Número                                    |
| B   | GRUPO MUSCULAR    | 22    | Dropdown                                  |
| C   | EJERCICIO         | 28    | Texto                                     |
| D   | SERIES            | 9     | Número entero                             |
| E   | REPS / TIEMPO     | 14    | Texto (ej: "12" o "45 seg")               |
| F   | PESO (kg)         | 10    | Número decimal / "Corporal"               |
| G   | DESCANSO (seg)    | 14    | Número entero                             |
| H   | DÍA(S)            | 22    | Texto (ej: "Lu-Mi-Vi")                    |
| I   | TÉCNICA / NOTAS   | 35    | Texto                                     |
| J   | COMPLETADO HOY    | 16    | Dropdown: ✓ Sí, ✗ No, - Día de descanso  |

Dropdown GRUPO MUSCULAR:
Pecho,Espalda,Hombros,Bíceps,Tríceps,Antebrazos,Abdomen/Core,Glúteos,
Cuádriceps,Isquiotibiales,Pantorrillas,Cardio,Movilidad/Flexibilidad,Cuerpo completo

### 10 ejercicios de ejemplo (filas 13-22)
Rutina full body 3 días (Lunes, Miércoles, Viernes):
1. Cuerpo completo, Sentadilla con barra, 4, 12, 60kg, 90, Lu-Mi-Vi, Rodillas no pasan los pies, -
2. Pecho, Press de banca plana, 3, 12, 50kg, 90, Lu-Mi-Vi, Bajar la barra al pecho, -
3. Espalda, Remo con barra, 3, 12, 40kg, 90, Lu-Mi-Vi, Codos pegados al cuerpo, -
4. Hombros, Press militar con mancuernas, 3, 12, 16kg, 75, Lu-Mi-Vi, No arquear la espalda, -
5. Cuádriceps, Prensa de piernas, 3, 15, 80kg, 90, Lu-Mi-Vi, Pies al ancho de hombros, -
6. Isquiotibiales, Curl femoral, 3, 12, 40kg, 75, Lu-Mi-Vi, Movimiento controlado, -
7. Bíceps, Curl con mancuernas, 3, 12, 12kg, 60, Lu-Vi, Sin balancear el cuerpo, -
8. Tríceps, Extensión en polea, 3, 15, 20kg, 60, Lu-Vi, Codos fijos, -
9. Abdomen/Core, Plancha frontal, 3, 45 seg, Corporal, 45, Lu-Mi-Vi, Cuerpo recto como tabla, -
10. Cardio, Caminata en cinta, 1, 20 min, Corporal, 0, Mi, Ritmo moderado 5.5km/h, -

### Tabla de seguimiento mensual (desde fila 25, columnas A:H)
Título "REGISTRO MENSUAL" en A25:H25, fondo #6C3483, texto blanco.
Tabla 4 semanas x 7 días:
- Fila encabezado: SEMANA | LUN | MAR | MIÉ | JUE | VIE | SÁB | DOM
- 4 filas de datos con Semana 1, 2, 3, 4
- Instrucción: "Escribe: E=Entrenado, D=Descanso, C=Cardio, X=No entrenado"
- Fórmula al final: =CONTAR.SI(B26:H29,"E")&" días entrenados este mes"

---

## HOJA 5: DIETA Y NUTRICIÓN

### Configuración
- Nombre pestaña: "🥗 DIETA"
- Color pestaña: #27AE60
- Sin gridlines

### Banner (filas 1-2)
Fondo #1E8449, texto: "CONTROL DE DIETA Y NUTRICIÓN"
Sub: "Una alimentación balanceada es la base de tu salud — registra, analiza y mejora"

### Panel de metas (filas 4-8, columnas A:F)
Título "MIS METAS DIARIAS" en A4:F4, fondo #27AE60, texto blanco.
Campos editables:
- Calorías objetivo: [número, ej: 2000] kcal/día
- Proteína objetivo: [número, ej: 120] g/día
- Carbohidratos objetivo: [número, ej: 250] g/día
- Grasas objetivo: [número, ej: 65] g/día
- Vasos de agua: [número, ej: 8] vasos/día (1 vaso = 250ml)

### Plan semanal de comidas (filas 10-18)
Header fila 10, fondo #27AE60, texto blanco, bold.
Columnas: DÍA | DESAYUNO | SNACK MAÑANA | ALMUERZO | SNACK TARDE | CENA | AGUA (vasos) | NOTA

Filas 11-17: Lunes a Domingo.
Fondo alternado: filas impares #F8F9FA, pares blanco.
Dropdown en col G (Agua): 1,2,3,4,5,6,7,8,9,10

Datos ejemplo (Lunes completo pre-llenado):
- Desayuno: Huevos revueltos 2u + pan francés 1u + café negro
- Snack AM: Plátano fresco 1 unidad
- Almuerzo: Arroz 1 taza + frijoles 1/2 taza + pollo a la plancha 150g + ensalada
- Snack PM: Yogur natural sin azúcar 150g
- Cena: Sopa de res + 2 tortillas de maíz
- Agua: 8
- Nota: Buen día - cumplí todas las comidas

### Tabla nutricional de referencia salvadoreña (columnas I:N, filas 4-28)
Título "GUÍA NUTRICIONAL SALVADOREÑA" en I4:N4, fondo #1A5276, texto blanco.
Header fila 5: ALIMENTO | PORCIÓN | CALORÍAS | PROT(g) | CARB(g) | GRASA(g)
Ancho columnas: I:24, J:14, K:11, L:10, M:10, N:10

20 alimentos típicos de El Salvador:
| Alimento                | Porción        | Cal | Prot | Carb | Grasa |
|-------------------------|----------------|-----|------|------|-------|
| Pupusa de queso         | 1 unidad (80g) | 230 | 8    | 35   | 7     |
| Pupusa de frijoles      | 1 unidad (80g) | 200 | 7    | 32   | 5     |
| Pupusa revuelta         | 1 unidad (80g) | 260 | 9    | 36   | 9     |
| Arroz blanco cocido     | 1 taza (186g)  | 206 | 4    | 45   | 0.4   |
| Frijoles rojos cocidos  | 1/2 taza (90g) | 113 | 7.5  | 20   | 0.5   |
| Pollo a la plancha      | 100g           | 165 | 31   | 0    | 3.6   |
| Carne de res asada      | 100g           | 250 | 26   | 0    | 16    |
| Huevo entero            | 1 unidad (60g) | 78  | 6    | 0.6  | 5     |
| Pan francés             | 1 unidad (40g) | 120 | 3.5  | 23   | 1.5   |
| Plátano fresco          | 1 unidad (120g)| 105 | 1.3  | 27   | 0.4   |
| Plátano frito maduro    | 1 unidad (100g)| 180 | 1    | 35   | 5     |
| Yuca cocida             | 100g           | 159 | 1.4  | 38   | 0.3   |
| Aguacate                | 1/4 de fruto   | 80  | 1    | 4    | 7     |
| Tortilla de maíz        | 1 unidad (25g) | 57  | 1.5  | 12   | 0.7   |
| Leche entera            | 1 vaso (240ml) | 149 | 8    | 11   | 8     |
| Queso duro seco         | 30g            | 110 | 7    | 0.5  | 9     |
| Atún en agua            | 1 lata (140g)  | 130 | 28   | 0    | 1     |
| Tomate mediano          | 1 unidad (130g)| 22  | 1    | 5    | 0.2   |
| Zanahoria               | 1 mediana (60g)| 25  | 0.6  | 6    | 0.1   |
| Naranja                 | 1 unidad (130g)| 62  | 1.2  | 15   | 0.2   |

---

## HOJA 6: HISTORIAL CLÍNICO

### Configuración
- Nombre pestaña: "📋 HISTORIAL"
- Color pestaña: #2C3E50
- Sin gridlines

### Banner (filas 1-2)
Fondo #2C3E50, texto: "HISTORIAL CLÍNICO PERSONAL"
Sub: "Documento de referencia médica — Comparte con tu médico en cada consulta"

### Sección 1: Datos médicos básicos (filas 4-14, columnas A:D)
Título "DATOS MÉDICOS BÁSICOS" en A4:D4, fondo #34495E, texto blanco.

| Fila | Etiqueta                  | Valor / Fórmula                                              |
|------|---------------------------|--------------------------------------------------------------|
| 5    | Grupo sanguíneo:          | Dropdown: O+,O-,A+,A-,B+,B-,AB+,AB-                        |
| 6    | Factor Rh:                | Dropdown: Positivo (+), Negativo (-)                         |
| 7    | Peso actual (kg):         | [número decimal]                                             |
| 8    | Talla (cm):               | [número entero]                                              |
| 9    | IMC:                      | =SI(D7="","",REDONDEAR(D7/((D8/100)^2),1))                  |
| 10   | Clasificación IMC (OMS):  | =SI(D9="","",SI(D9<18.5,"Bajo peso",SI(D9<25,"Normal",SI(D9<30,"Sobrepeso","Obesidad")))) |
| 11   | Presión arterial habitual:| [texto, ej: 120/80 mmHg]                                    |
| 12   | Frecuencia cardíaca:      | [número] lpm en reposo                                       |
| 13   | Glucosa en ayunas:        | [número] mg/dL                                               |
| 14   | Alergias medicamentosas:  | [texto libre]                                                |

Formato condicional en IMC (D9):
- Azul claro (#D6EAF8): < 18.5 (Bajo peso)
- Verde (#D5F5E3): 18.5-24.9 (Normal)
- Amarillo (#FEF9E7): 25-29.9 (Sobrepeso)
- Rojo claro (#FADBD8): >= 30 (Obesidad)

### Sección 2: Seguimiento de signos vitales (filas 16-26)
Título "SEGUIMIENTO DE SIGNOS VITALES" en A16:G16, fondo #1A5276, texto blanco.
Header fila 17: FECHA | PESO (kg) | PRESIÓN SISTÓLICA | PRESIÓN DIASTÓLICA | GLUCOSA (mg/dL) | FREC. CARDÍACA | NOTAS

5 registros de ejemplo con fechas de los últimos 5 meses:
- Variación realista: peso entre 72-74kg, presión 125/82 - 120/80, glucosa 95-115

### Sección 3: Vacunas (filas 28-36)
Título "REGISTRO DE VACUNAS" en A28:E28, fondo #117A65, texto blanco.
Header: VACUNA | DOSIS | FECHA APLICACIÓN | PRÓXIMA DOSIS | LUGAR DE APLICACIÓN

Datos ejemplo (vacunas comunes en El Salvador):
- COVID-19, Completa (2 dosis), 15/08/2021, No requerida/Refuerzo según indicación, ISSS
- Influenza, Anual, 01/04/2026, 01/04/2027, Clínica comunal
- Tétanos (DT), Refuerzo, 10/02/2023, 10/02/2033, Hospital Nacional
- Hepatitis B, Completa (3 dosis), 2005, No requerida, MSPAS

### Sección 4: Antecedentes quirúrgicos (filas 38-44)
Título "CIRUGÍAS Y PROCEDIMIENTOS" en A38:F38, fondo #6C3483, texto blanco.
Header: FECHA | PROCEDIMIENTO | CIRUJANO | HOSPITAL | RESULTADO | NOTAS

### Sección 5: Antecedentes familiares (filas 46-55)
Título "ANTECEDENTES FAMILIARES" en A46:D46, fondo #922B21, texto blanco.
Header: FAMILIAR | CONDICIÓN | EDAD DIAGNÓSTICO | OBSERVACIONES

Familiares predefinidos (solo etiquetas, datos vacíos para que el usuario llene):
Padre, Madre, Abuelo paterno, Abuela paterna, Abuelo materno, Abuela materna, Hermano/a 1

### Sección 6: Médicos tratantes (filas 57-65)
Título "MIS MÉDICOS TRATANTES" en A57:F57, fondo #0F4C8A, texto blanco.
Header: ESPECIALIDAD | NOMBRE COMPLETO | CLÍNICA / HOSPITAL | TELÉFONO | CORREO | FREC. VISITA

Dropdown FREC. VISITA: Mensual,Cada 3 meses,Cada 6 meses,Anual,Según necesidad,De urgencias

---

## HOJA 7: INSTRUCCIONES DE USO

### Configuración
- Nombre pestaña: "📖 INSTRUCCIONES"
- Color pestaña: #7F8C8D
- Sin gridlines

### Banner principal (filas 1-3, A:H)
Merge A1:H3, fondo #2C3E50
Texto: "¡Bienvenido a SaludSV Pro!" — fuente 18pt, Bold, Blanco
Sub: "Tu salud, organizada. Tu vida, protegida." — fuente 12pt, Italic, #A9DFBF

### Sección 1: Cómo usar cada módulo (filas 5-35)
Para cada sección del archivo, crear un bloque visual con:
- Título con emoji y color de esa pestaña
- Descripción en 2-3 líneas
- Lista de acciones que puede hacer el usuario (3-4 puntos)

Orden: INICIO → CITAS → MEDICAMENTOS → EJERCICIOS → DIETA → HISTORIAL

### Sección 2: Guía de colores (filas 37-50)
Tabla visual explicando cada color con rectángulos de color como ejemplo:
| Color visual | Significado | Módulo donde aparece |
|--- (implementar con fondo de celda como muestra del color)

### Sección 3: Cómo importar citas al calendario (filas 52-62)
Título: "CÓMO AGREGAR CITAS A GOOGLE CALENDAR"
Pasos numerados:
1. Abre la carpeta "calendario/" que se generó junto a este archivo
2. Busca el archivo .ics de tu cita (ej: cita_001.ics)
3. Haz doble clic en el archivo — Google Calendar se abrirá automáticamente
4. Confirma la importación — el evento incluye 2 recordatorios automáticos
5. Recordatorios configurados: 1 día antes y 2 horas antes de la cita

### Sección 4: Recomendaciones (filas 64-72)
Caja con fondo #EBF5FB, borde #2E86C1:
"RECOMENDACIONES DE USO:
• Actualiza tus datos cada semana (preferiblemente el domingo)
• Lleva el archivo en una USB o Google Drive para acceder desde cualquier PC
• Comparte el historial clínico con tu médico en cada consulta
• Haz una copia de seguridad con la fecha en el nombre: SaludSV_Pro_[MES_AÑO].xlsx
• Cambia tu rutina de ejercicios cada 6 semanas para mejores resultados"

### Pie de página (filas 74-76)
Texto centrado:
"SaludSV Pro v1.0 | Desarrollado con Claude AI por [Tu Nombre] | El Salvador 2026"
"Para más información: generado como proyecto del Bootcamp de Claude | Anthropic"

---

## GENERACIÓN DE ARCHIVOS .ICS

### Para cada cita médica de ejemplo, generar un archivo .ics

Usar la librería 'icalendar' con la siguiente estructura:

```python
from icalendar import Calendar, Event, Alarm
from datetime import datetime, timedelta
import pytz

def generar_ics(numero, fecha, hora, doctor, especialidad, hospital, direccion, motivo, output_dir):
    """
    Genera un archivo .ics para importar a Google Calendar o Outlook.
    Incluye 2 alarmas: 1 día antes y 2 horas antes.
    """
    tz = pytz.timezone('America/El_Salvador')
    
    # Crear calendario
    cal = Calendar()
    cal.add('prodid', '-//SaludSV Pro//El Salvador//ES')
    cal.add('version', '2.0')
    cal.add('calscale', 'GREGORIAN')
    cal.add('method', 'PUBLISH')
    
    # Crear evento
    event = Event()
    dt_inicio = tz.localize(datetime.combine(fecha, hora))
    dt_fin = dt_inicio + timedelta(hours=1)
    
    event.add('summary', f'Cita: {especialidad} - {doctor}')
    event.add('dtstart', dt_inicio)
    event.add('dtend', dt_fin)
    event.add('location', f'{hospital} - {direccion}')
    event.add('description', f'Motivo: {motivo}\nMédico: {doctor}\nHospital: {hospital}')
    
    # Alarma 1: 1 día antes
    alarm1 = Alarm()
    alarm1.add('action', 'DISPLAY')
    alarm1.add('description', f'Recordatorio: Mañana tienes cita con {doctor}')
    alarm1.add('trigger', timedelta(days=-1))
    event.add_component(alarm1)
    
    # Alarma 2: 2 horas antes
    alarm2 = Alarm()
    alarm2.add('action', 'DISPLAY')
    alarm2.add('description', f'Tu cita con {doctor} es en 2 horas - {hospital}')
    alarm2.add('trigger', timedelta(hours=-2))
    event.add_component(alarm2)
    
    cal.add_component(event)
    
    # Guardar archivo
    filename = os.path.join(output_dir, f'cita_{numero:03d}.ics')
    with open(filename, 'wb') as f:
        f.write(cal.to_ical())
    
    return filename
```

Generar un .ics para cada una de las 3 citas de ejemplo.

---

## FUNCIONES AUXILIARES REQUERIDAS EN EL SCRIPT

```python
# Importaciones necesarias
import os
import sys
from datetime import datetime, date, timedelta
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side,
    GradientFill
)
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from icalendar import Calendar, Event, Alarm
import pytz

def make_fill(hex_color):
    """Crea un PatternFill sólido desde un hex color (sin #)"""
    return PatternFill(start_color=hex_color.replace('#',''), 
                       end_color=hex_color.replace('#',''), 
                       fill_type='solid')

def make_font(size=11, bold=False, color='2C3E50', italic=False, name='Calibri'):
    """Crea una fuente con los parámetros dados"""
    return Font(name=name, size=size, bold=bold, 
                color=color.replace('#',''), italic=italic)

def make_border(style='thin', color='BDC3C7'):
    """Crea un borde completo con el estilo dado"""
    side = Side(style=style, color=color.replace('#',''))
    return Border(left=side, right=side, top=side, bottom=side)

def make_alignment(horizontal='left', vertical='center', wrap=False):
    """Crea un alineamiento"""
    return Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap)

def apply_header_style(ws, row, col_start, col_end, text, bg_color, 
                        font_size=12, merge=True):
    """Aplica estilo de header a una fila, con merge opcional"""
    if merge:
        ws.merge_cells(start_row=row, start_column=col_start,
                       end_row=row, end_column=col_end)
    cell = ws.cell(row=row, column=col_start, value=text)
    cell.fill = make_fill(bg_color)
    cell.font = make_font(size=font_size, bold=True, color='FFFFFF')
    cell.alignment = make_alignment(horizontal='center')
    cell.border = make_border(color='FFFFFF')
    ws.row_dimensions[row].height = 28

def apply_col_header(ws, row, headers_config):
    """
    headers_config: lista de dicts con keys: col, text, width, bg_color
    """
    for h in headers_config:
        cell = ws.cell(row=row, column=h['col'], value=h['text'])
        cell.fill = make_fill(h.get('bg_color', '0F4C8A'))
        cell.font = make_font(size=11, bold=True, color='FFFFFF')
        cell.alignment = make_alignment(horizontal='center')
        cell.border = make_border(color='FFFFFF')
        if 'width' in h:
            ws.column_dimensions[get_column_letter(h['col'])].width = h['width']
    ws.row_dimensions[row].height = 30

def add_dropdown(ws, col_letter, row_start, row_end, options_list):
    """Agrega un dropdown de validación de datos"""
    formula = '"' + ','.join(options_list) + '"'
    dv = DataValidation(type='list', formula1=formula, allow_blank=True)
    dv.sqref = f"{col_letter}{row_start}:{col_letter}{row_end}"
    ws.add_data_validation(dv)

def set_banner(ws, text, subtext, color_hex, rows_merge='A1:H2', height=55):
    """Crea el banner principal de cada hoja"""
    # Merge cells, apply fill, set text
    ws.merge_cells(rows_merge)
    cell = ws[rows_merge.split(':')[0]]
    cell.value = text + '\n' + subtext
    cell.fill = make_fill(color_hex)
    cell.font = make_font(size=16, bold=True, color='FFFFFF', name='Calibri')
    cell.alignment = Alignment(horizontal='center', vertical='center', 
                                wrap_text=True)
    # Adjust row heights
    first_row = int(rows_merge.split(':')[0][1:])
    last_row = int(rows_merge.split(':')[1][1:])
    total_height = height
    per_row = total_height / (last_row - first_row + 1)
    for r in range(first_row, last_row + 1):
        ws.row_dimensions[r].height = per_row
```

---

## FUNCIÓN MAIN Y FLUJO DE EJECUCIÓN

```python
def main():
    print("\n" + "="*62)
    print("   SaludSV Pro — Generador de Organizador de Salud")  
    print("   Versión 1.0 | El Salvador 2026 | Powered by Claude AI")
    print("="*62 + "\n")
    
    # 1. Crear estructura de carpetas
    output_dir = os.path.join(os.getcwd(), 'output')
    cal_dir = os.path.join(output_dir, 'calendario')
    os.makedirs(cal_dir, exist_ok=True)
    print("✓ Carpetas de output creadas")
    
    # 2. Crear workbook
    wb = Workbook()
    wb.remove(wb.active)  # Remover hoja default
    
    # 3. Crear cada hoja
    print("→ Creando hoja: INICIO...")
    crear_hoja_inicio(wb)
    
    print("→ Creando hoja: CITAS MÉDICAS...")
    citas = crear_hoja_citas(wb)
    
    print("→ Creando hoja: MEDICAMENTOS...")
    crear_hoja_medicamentos(wb)
    
    print("→ Creando hoja: EJERCICIOS...")
    crear_hoja_ejercicios(wb)
    
    print("→ Creando hoja: DIETA...")
    crear_hoja_dieta(wb)
    
    print("→ Creando hoja: HISTORIAL CLÍNICO...")
    crear_hoja_historial(wb)
    
    print("→ Creando hoja: INSTRUCCIONES...")
    crear_hoja_instrucciones(wb)
    
    # 4. Activar hoja INICIO al abrir
    wb.active = wb["🏥 INICIO"]
    
    # 5. Guardar Excel
    excel_path = os.path.join(output_dir, 'SaludSV_Pro.xlsx')
    wb.save(excel_path)
    print(f"\n✓ Excel generado: {excel_path}")
    
    # 6. Generar .ics para cada cita
    print("\n→ Generando archivos de calendario (.ics)...")
    for i, cita in enumerate(citas, 1):
        ics_path = generar_ics(i, cita_dir=cal_dir, **cita)
        print(f"  ✓ {os.path.basename(ics_path)}")
    
    # 7. Generar README.md
    generar_readme(output_dir)
    print("✓ README.md generado")
    
    print("\n" + "="*62)
    print("  ¡SaludSV Pro generado exitosamente!")
    print(f"  Archivos en: {output_dir}/")
    print("  → SaludSV_Pro.xlsx (abre en Excel o LibreOffice)")
    print(f"  → calendario/ ({len(citas)} archivos .ics)")
    print("  → README.md")
    print("="*62 + "\n")

if __name__ == '__main__':
    main()
```

---

## CONSIDERACIONES TÉCNICAS IMPORTANTES

### Compatibilidad
- El archivo debe abrir correctamente en Excel 2016+, Excel 365, LibreOffice Calc 7+
- NO usar macros VBA — solo features nativos de Excel
- Los formatos condicionales deben funcionar con versiones antiguas de openpyxl

### Zona horaria y localización
- Zona horaria para .ics: America/El_Salvador (UTC-6)
- Formato de fecha en celdas: DD/MM/AAAA
- Separador decimal: punto (.)
- Moneda: dólares USD ($) — El Salvador usa dólar estadounidense

### Formato condicional con openpyxl
Usar `ConditionalFormattingList` y `FormulaRule` para colorear filas completas.
Ejemplo para colorear fila entera basado en valor de columna I:
```python
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import FormulaRule

red_fill = make_fill('#FADBD8')
red_font = make_font(color='#922B21')
dxf_red = DifferentialStyle(fill=red_fill, font=red_font)

# Aplicar a filas 5:104 cuando col I = "Cancelada"
ws.conditional_formatting.add(
    'A5:K104',
    FormulaRule(formula=['$I5="Cancelada"'], dxf=dxf_red)
)
```

### DataValidation con listas largas
Si la lista de opciones supera 255 caracteres, crear una hoja auxiliar oculta
con los valores y referenciar esa hoja en la validación:
```python
# Crear hoja auxiliar oculta
ws_aux = wb.create_sheet('_aux_listas')
ws_aux.sheet_state = 'hidden'

# Escribir lista en columna A
especialidades = ['Medicina General', 'Cardiología', ...]
for i, esp in enumerate(especialidades, 1):
    ws_aux.cell(row=i, column=1, value=esp)

# Referenciar en DataValidation
dv = DataValidation(
    type='list',
    formula1=f"_aux_listas!$A$1:$A${len(especialidades)}"
)
```

### Manejo de errores
Envolver cada función `crear_hoja_*` en try/except con mensaje descriptivo:
```python
try:
    crear_hoja_citas(wb)
    print("✓ Hoja CITAS creada")
except Exception as e:
    print(f"✗ Error en hoja CITAS: {e}")
    sys.exit(1)
```

---

## CALIDAD Y PRUEBAS

El script debe cumplir estos criterios antes de considerarse completo:
1. Ejecutarse con `python salud_sv_pro.py` sin errores ni warnings
2. Generar el archivo en menos de 30 segundos
3. El Excel no debe mostrar errores de fórmula (#REF!, #N/A, #VALUE!, #DIV/0!)
4. Todos los dropdowns deben funcionar al hacer clic en las celdas
5. Los formatos condicionales deben activarse visualmente con los datos de ejemplo
6. Los 3 archivos .ics deben importarse correctamente a Google Calendar
7. El archivo Excel debe pesar menos de 2MB
8. Las pestañas deben tener los colores correctos asignados
9. La hoja que se ve al abrir el archivo debe ser "🏥 INICIO"
10. Imprimir progreso en consola durante la generación

---

## GENERACIÓN DEL README.md

Generar un archivo README.md en la carpeta output/ con este contenido:

```markdown
# SaludSV Pro — Organizador Personal de Salud

**Versión:** 1.0  
**País:** El Salvador  
**Tecnología:** Python + openpyxl + icalendar

## ¿Qué incluye?

- `SaludSV_Pro.xlsx` — Tu organizador de salud personal con 7 secciones
- `calendario/` — Archivos .ics para importar tus citas a Google Calendar

## Cómo usar el archivo Excel

1. Abre `SaludSV_Pro.xlsx` en Excel (2016+) o LibreOffice Calc
2. Comienza en la hoja **🏥 INICIO** y llena tus datos personales
3. Registra tus citas en **📅 CITAS MÉDICAS**
4. Agrega tus medicamentos en **💊 MEDICAMENTOS**
5. Configura tu rutina en **💪 EJERCICIOS**
6. Planifica tu dieta en **🥗 DIETA**
7. Completa tu historia clínica en **📋 HISTORIAL**

## Cómo importar citas al calendario

1. Ve a la carpeta `calendario/`
2. Haz doble clic en `cita_001.ics`
3. Se abrirá Google Calendar o Outlook para confirmar
4. El evento incluye recordatorios automáticos: 1 día antes y 2 horas antes

## Instalación (para regenerar el archivo)

```bash
pip install -r requirements.txt
python salud_sv_pro.py
```

## Créditos

Desarrollado como proyecto del Bootcamp de Claude | Anthropic  
El Salvador, cubo ai 2026
```

---

*Fin de especificaciones — SaludSV Pro v1.0*
*Este archivo CLAUDE.md es la fuente de verdad del proyecto. Seguir todas las especificaciones al pie de la letra.*
