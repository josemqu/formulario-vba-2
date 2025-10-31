# 🧩 Esquema Técnico de Base de Datos – Sistema de Registro y Análisis de Incidentes de Tránsito

## 📘 Descripción General

El presente esquema de base de datos define la estructura técnica para el registro, análisis y procesamiento de incidentes de tránsito.  
Está diseñado para permitir el almacenamiento estructurado de información sobre **personas involucradas**, **vehículos** y el propio **incidente**, con el objetivo de realizar análisis estadístico y generar reportes dinámicos.

---

## 🏗️ Estructura General

El modelo relacional está compuesto por las siguientes tablas principales:

- **Incidente** → Contiene la información general del suceso.
- **Persona** → Registra los datos de las personas afectadas (empleados YPF, contratistas o terceros).
- **Vehiculo** → Detalla las características y condiciones de los vehículos involucrados.

Cada **Incidente** puede tener múltiples **Personas** y **Vehículos** asociados.

---

## 🔗 Relaciones Entre Tablas

| Tabla Origen | Tabla Destino | Tipo de Relación | Descripción |
|---------------|----------------|------------------|--------------|
| `Incidente` | `Persona` | 1 a N | Un incidente puede involucrar una o más personas. |
| `Incidente` | `Vehiculo` | 1 a N | Un incidente puede involucrar uno o más vehículos. |

---

## 🧱 Tablas y Campos

### 🟦 Tabla: `Incidente`
| Campo | Tipo | Descripción | Celda del formulario |
|--------|------|-------------|----------------------|
| `id_incidente` | Integer (PK, autoincremental) | Identificador único del incidente | Form!D5 |
| `fecha_ocurrencia` | Date | Fecha del hecho | Form!D6 |
| `hora_ocurrencia` | Time | Hora del hecho | Form!D7 |
| `pais` | Texto | País donde ocurrió el incidente | Form!D8 |
| `provincia` | Texto | Provincia (según catálogo YPF GestiónCAS) | Form!D9 |
| `localidad_zona` | Texto | Localidad o zona (según catálogo YPF GestiónCAS) | Form!D10 |
| `coordenadas_geograficas` | Texto | Coordenadas GPS (lat,lng) | Form!D11 |
| `lugar_especifico` | Texto | Descripción del lugar (máx. 100 caracteres) | Form!D12 |
| `uo_incidente` | Texto | Unidad Operativa donde ocurrió el hecho | Form!D13 |
| `uo_accidentado` | Texto | Unidad Operativa a la que pertenece el accidentado | Form!D14 |
| `tipo_consecuencia` | Texto | Tipo de consecuencia | Form!D15 |
| `material_sustancia_peligrosa` | Texto | Material o sustancia peligrosa | Form!D16 |
| `descripcion_esv` | Texto largo | Descripción narrativa del evento | Form!D17 |
| `denuncia_policial` | Enum(SI, NO, NA) | Indica si hubo denuncia policial | Form!D22 |
| `lugar_denuncia_policial` | Texto | Lugar de la denuncia policial | Form!D23 |
| `examen_alcoholemia` | Enum(SI, NO, NA) | Resultado o existencia del examen | Form!D24 |
| `examen_sustancias` | Enum(SI, NO, NA) | Resultado o existencia del examen | Form!D25 |
| `entrevistas_testigos` | Enum(SI, NO, NA) | Indica si hubo entrevistas con testigos | Form!D26 |
| `accion_inmediata` | Texto | Breve descripción de las acciones inmediatas | Form!D27 |
| `consecuencias_seguridad` | Enum(SI, NO) | Indica si hubo consecuencias reales o potenciales | Form!D28 |
| `fecha_reporte` | Date | Fecha del reporte | Form!D29 |
| `cantidad_personas` | Integer | Número de personas afectadas | Form!D30 |
| `cantidad_vehiculos` | Integer | Número de vehículos involucrados | Form!D31 |
| `clase_evento` | Texto | Tipo o clase del evento (catálogo YPF) | Form!D32 |
| `tipo_colision` | Texto | Tipo de colisión (frontal, lateral, etc.) | Form!D33 |
| `nivel_severidad` | Enum(S0-S5) | Nivel de severidad según clasificación YPF | Form!D34 |
| `clasificacion_esv` | Enum(Severo, Potencialmente Severo, Menor) | Clasificación del evento | Form!D35 |
| `tipo_superficie` | Texto | Tipo de superficie (asfalto, ripio, etc.) | Form!AC6
| `posee_banquina` | Enum(SI, NO, NA) | Indica si posee banquina | Form!AC7
| `tipo_ruta` | Texto | Tipo de vía | Form!AC8
| `velocidad_max_permitida_YPF` | Texto | Velocidad máxima permitida por YPF (10-30Km/h, 31-40km/h, 41-60 Km/h, 61-80 Km/h,81-100 Km/h, >100Km/h) | Form!AC9
| `densidad_trafico` | Enum(Alta, Media, Baja) | Tráfico al momento del hecho | Form!AC10
| `condicion_ruta` | Texto | Estado general del camino | Form!AC11
| `iluminacion_ruta` | Texto | Condición de luz natural o artificial | Form!AC12
| `senalizacion_ruta` | Texto | Estado de la señalización | Form!AC13
| `geometria_ruta` | Texto | Curvatura o pendiente | Form!AC14
| `condiciones_climaticas` | Texto | Condiciones meteorológicas (Seco y templado, Lluvioso, Tormenta, Niebla, Humo sobre la ruta, Calor Extremo, Granizo, Hielo, Viento fuerte) | Form!AC15
| `rango_temperaturas` | Texto | Rango de temperatura ambiental | Form!AC16

---

### 🟩 Tabla: `Persona`
| Campo | Tipo | Descripción | Celda del formulario |
|--------|------|-------------|----------------------|
| `id_persona` | Integer (PK) | Identificador único de persona | Form!K5:T5 |
| `id_incidente` | Integer (FK → Incidente.id_incidente) | Relación con incidente | Form!D5 |
| `nombre_persona` | Texto | Nombre | Form!K6:T6 |
| `apellido_persona` | Texto | Apellido | Form!K7:T7 |
| `edad_persona` | Integer | Edad (0-100) | Form!K8:T8 |
| `tipo_persona` | Enum(YPF, Contratista, Tercero) | Tipo de involucrado | Form!K9:T9 |
| `rol_persona` | Enum(Conductor, Acompañante, Operador, Otro) | Rol dentro del incidente | Form!K10:T10 |
| `antiguedad_persona` | Texto | Rango de antigüedad o "No aplica" | Form!K11:T11 |
| `tipo_tarea` | Enum(Rutinaria, Especial, Emergencia, Supervisión, Apoyo Logístico, NA) | Tipo de tarea | Form!K12:T12 |
| `turno_operativo` | Enum(Diurno, Nocturno, Extendido-Mixto, NA) | Turno operativo | Form!K13:T13 |
| `tipo_danio_persona` | Enum(Fatalidad, Accidente, Primeros Auxilios, Ninguno) | Tipo de daño | Form!K14:T14 |
| `dias_perdidos` | Integer | Días de ausencia | Form!K15:T15 |
| `atencion_medica` | Enum(SI, NO, NA) | Si recibió atención médica | Form!K16:T16 |
| `in_itinere` | Enum(SI, NO, NA) | Indica si ocurrió en itinere | Form!K17:T17 |
| `tipo_afectacion` | Texto | Tipo de afectación física | Form!K18:T18 |
| `parte_afectada` | Texto | Parte del cuerpo afectada | Form!K19:T19 |
| `clase_licencia` | Texto | Clase de licencia (Automóviles particulares, Camiones sin acoplado, Transporte de pasajeros hasta 8 asientos, Camión con acoplado, Maquinaria especial/carga peligrosa (segun pais), Transporte de personas con discapacidad) | Form!K20:T20 |
| `entrenamiento` | Texto | Indica que entrenamiento posee (Seleccion multiple de: Manejo defensivo, Gestión de Fatiga, Conducción en ripio, Normativa Corporativa, Utilización del teléfono movil al conducir, Campaña de efectos de Alcohol y drogas al conducir, Política de conducción nocturna, Velocidades maximas YPF en caso de condiciones climáticas adversas (lluvia, nieve, viento, tormenta,etc), Velocidad maxima YPF al ingreso/egreso de rotonda, Manejo de crisis en caso de accidentes en la via pública, Otros (se debe poder escribir ,y que no figuren en la lista previa)) |  Form!K21:T21 |
| `aptitud_tarea` | Texto | Indica que aptitud posee (Apto, Apto con restricciones y/o tratamiento médico aprobado, No apto) |  Form!K22:T22 |

---

### 🟨 Tabla: `Vehiculo`

| Campo | Tipo | Descripción | Celda del formulario |
|-------|------|-------------|----------------------|
| `id_vehiculo` | Integer (PK) | Identificador único del vehículo (PK). | Form!W5:Z5 |
| `id_incidente` | Integer (FK → Incidente.id_incidente) | FK a `Incidente.id_incidente`. Vincula el vehículo con su incidente. | Form!D5 |
| `tipo_vehiculo` | Texto | Tipo de vehículo. Opciones: Bicicleta, Moto, Ciclomotor, Automóvil, Auto utilitario, Minibús, Ómnibus, Pickup, Camión chasis, Camión con Cisterna, Camión Pluma, Camión Volcador, Motoniveladora, Retroexcavadora, Pala cargadora, Topadora, Grúa, Trailer, Side-by-Side/UTV, Vehículo adaptado (discapacidad). | Form!W6:Z6 |
| `duenio_vehiculo` | Texto | Dueño del vehículo. Opciones: Propio, Contratista, Tercero. | Form!W7:Z7 |
| `uso_vehiculo` | Texto | Uso del vehículo. Opciones: Comercial, Particular, Otro, No se sabe. | Form!W8:Z8 |
| `posee_patente` | Enum(SI, NO, NA) | Indicador de patente. Opciones: SI / NO. | Form!W9:Z9 |
| `numero_patente` | Texto | Patente alfanumérica; valores especiales: "desconocida" / "NA". | Form!W10:Z10 |
| `anio_fabricacion_vehiculo` | Texto | Año de fabricación o texto libre (ej. "desconocido"). | Form!W11:Z11 |
| `tarea_vehiculo` | Texto | Tarea que realizaba. Ej.: transporte de personas, transporte de cargas generales, transporte de sustancias peligrosas/combustible, maniobrando/estacionando, trayecto de regreso vacío, transporte de mercaderías, tareas generales, viaje familiar/comercial, se desconoce, otros. | Form!W12:Z12 |
| `estado_vehiculo` | Texto | Estado del vehículo. Opciones: Bueno, Regular, Deficiente | Form!W13:Z13 |
| `tipo_danio_vehiculo` | Texto | Clasificación del daño. Opciones: Destrucción total; Daños en carrocería que afectan continuidad de viaje (remolque/grúa); Daños en carrocería que NO afectan continuidad de viaje; Daños mecánicos/eléctricos que afectan continuidad; Daños mecánicos/eléctricos que no afectan continuidad; Daños leves; Sin daños. | Form!W14:Z14 |
| `cinturon_seguridad` | Enum(SI, NO, NA) | Uso de cinturón. Opciones: SI / NO. | Form!W15:Z15 |
| `cabina_cuchetas` | Enum(SI, NO, NA) | Cabina con cuchetas. Opciones: SI / NO. | Form!W16:Z16 |
| `airbags` | Enum(SI, NO, NA) | Airbags presentes/desplegados. Opciones: SI / NO. | Form!W17:Z17 |
| `gestion_flotas` | Enum(SI, NO, NA) | Gestion de flotas. Opciones: SI / NO. | Form!W18:Z18 |
| `token_conductor` | Enum(SI, NO, NA) | Token del conductor. Opciones: SI / NO / NA. | Form!W19:Z19 |
| `marca_dispositivo` | Texto | Marca de dispositivo telemático. Opciones: Microtrack, ITURAN, IMSEG, Otros. | Form!W20:Z20 |
| `deteccion_fatiga` | Enum(SI, NO, NA) | Sistema detección fatiga/distracción. Opciones: SI / NO / NA. | Form!W21:Z21 |
| `limitador_velocidad` | Enum(SI, NO, NA) | Sistema limitador de velocidad. Opciones: SI / NO / NA. | Form!W22:Z22 |
| `camara_trasera` | Enum(SI, NO, NA) | Cámara trasera. Opciones: SI / NO / NA. | Form!W23:Z23 |
| `camara_delantera` | Enum(SI, NO, NA) | Cámara delantera. Opciones: SI / NO / NA. | Form!W24:Z24 |
| `camara_punto_ciego` | Enum(SI, NO, NA) | Cámaras en puntos ciegos. Opciones: SI / NO / NA. | Form!W25:Z25 |
| `camara_360` | Enum(SI, NO, NA) | Cámara 360°. Opciones: SI / NO / NA. | Form!W26:Z26 |
| `espejo_punto_ciego` | Enum(SI, NO, NA) | Espejos para puntos ciegos. Opciones: SI / NO / NA. | Form!W27:Z27 |
| `alarma_marcha_atras` | Enum(SI, NO, NA) | Alarma de marcha atrás. Opciones: SI / NO / NA. | Form!W28:Z28 |
| `sistema_frenos` | Texto | Tipo de sistema de frenos. Opciones: Sistema de frenos antibloqueo, Frenado electrónico, Estabilidad electrónica, Soporte anti-vuelco, Frenado de emergencia avanzado, Alerta de frenado de emergencia. | Form!W29:Z29 |
| `monitoreo_neumaticos` | Enum(SI, NO, NA) | Monitoreo de presión de neumáticos. Opciones: SI / NO / NA. | Form!W30:Z30 |
| `proteccion_lateral` | Enum(SI, NO, NA) | Protección lateral para bicicletas/motos. Opciones: SI / NO / NA. | Form!W31:Z31 |
| `proteccion_trasera` | Enum(SI, NO, NA) | Protección trasera antiempotramiento. Opciones: SI / NO / NA. | Form!W32:Z32 |
| `acondicionador_cabina` | Enum(SI, NO, NA) | Aire acondicionado en cabina. Opciones: SI / NO / NA. | Form!W33:Z33 |
| `calefaccion_cabina` | Enum(SI, NO, NA) | Calefacción en cabina. Opciones: SI / NO / NA. | Form!W34:Z34 |
| `manos_libres_cabina` | Enum(NO POSEE / FUNCIONA / NO FUNCIONA / DESHABILITADO) | Sistema Bluetooth/manos libres. Opciones: NO POSEE / FUNCIONA / NO FUNCIONA / DESHABILITADO. | Form!W35:Z35 |
| `kit_alcoholemia` | Enum(En cabina / Alcoholímetro antiarranque / No posee) | Control alcoholemia en cabina. Opciones: En cabina / Alcoholímetro antiarranque / No posee. | Form!W36:Z36 |
| `kit_emergencia` | Texto | Kit de emergencia presente. Opciones múltiples: Primeros auxilios, Chaleco reflectivo, Conos/triángulos, Matafuegos, Kit para derrames, No Tiene, Incompleto. | Form!W37:Z37 |
| `epps_vehiculo` | Texto | EPPs disponibles para conductor/acompañante. Opciones múltiples: Botines seguridad, Casco y máscara/anteojos, Ropa ignífuga (transp. sustancias peligrosas), Guantes de descarga, Guantes auxilio mecánico. | Form!W38:Z38 |
| `observaciones_vehiculo` | Texto | Campo libre para observaciones adicionales (daños, nota técnica, referencia a evidencias/ fotos). | Form!W39:Z39 |

