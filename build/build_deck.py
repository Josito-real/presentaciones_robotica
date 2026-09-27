#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Construye 'Actuadores - rediseño.pptx' a partir del material de NotebookLM."""
import os
import sys
sys.path.insert(0, os.path.dirname(__file__))
from design import *  # noqa
from design import _set_run
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

ROOT = os.path.join(os.path.dirname(__file__), "..")
D = os.path.join(ROOT, "imagenes", "diagramas")
F = os.path.join(ROOT, "imagenes", "fotos")
G = os.path.join(ROOT, "imagenes", "graficas")

prs = new_presentation()


def _arrow(conn):
    line = conn.line._get_or_add_ln()
    tail = line.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'})
    line.append(tail)

# ====================================================================
# 1. PORTADA
# ====================================================================
s = blank_slide(prs, bg=NAVY)
_, tf = add_textbox(s, Inches(0.75), Inches(0.6), Inches(5.0), Inches(0.35))
p = tf.paragraphs[0]
r = p.add_run(); _set_run(r, "UNIMET  ·  ROBÓTICA INDUSTRIAL", size=13, color=ORANGE, bold=True, spacing=1.8)

_, tf = add_textbox(s, Inches(0.75), Inches(2.5), Inches(6.9), Inches(3.2))
p = tf.paragraphs[0]; set_paragraph(p, line_spacing=1.02)
r = p.add_run(); _set_run(r, "Solenoides y Motores DC:", font=FONT_TITLE, size=40, color=WHITE, bold=True)
p2 = tf.add_paragraph(); set_paragraph(p2, line_spacing=1.02)
r2 = p2.add_run(); _set_run(r2, "de la Lógica al Movimiento", font=FONT_TITLE, size=40, color=ORANGE, bold=True)
p3 = tf.add_paragraph(); set_paragraph(p3, space_before=22, line_spacing=1.3)
r3 = p3.add_run(); _set_run(r3, "Principios de funcionamiento, modelado físico e interfaces\nde control para Robótica Industrial.",
                              font=FONT_BODY, size=16, color=RGBColor(0xC9, 0xD3, 0xE0))

rounded_image_frame(s, f"{D}/motor_exploded_portada.png", Inches(7.9), Inches(0), SLIDE_W - Inches(7.9), SLIDE_H, cover=True)
add_notes(s, "Ilustración: vista explotada de un motor DC con reductora, tomada del material original del curso.")

# ====================================================================
# 2. CLASIFICACIÓN DE ACTUADORES
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Fundamentos", "Clasificación de actuadores electromecánicos")
add_picture_contain(s, f"{D}/clasificacion_actuadores.png", Inches(0.6), Inches(1.9), Inches(12.1), Inches(4.6))
_, tf = add_textbox(s, Inches(0.6), Inches(6.6), Inches(12.1), Inches(0.5))
p = tf.paragraphs[0]
r = p.add_run(); _set_run(r, "Los actuadores convierten energía eléctrica en movimiento: lineal (solenoides, actuadores lineales) o rotativo (motor DC, paso a paso, servomotor).",
                            size=13.5, color=GRAY_TXT, italic=True)
add_page_footer(s, 2)

# ====================================================================
# 3. DIVISORA 01 — SOLENOIDES Y RELÉS
# ====================================================================
add_divider_slide(prs, "01", "Solenoides y relés",
                   "Del pulso digital a la fuerza mecánica: la forma más simple de actuador electromagnético.",
                   f"{F}/rele_partes.jpg",
                   credit="Foto: David Boettcher — dominio público (Wikimedia Commons)")

# ====================================================================
# 4. SOLENOIDE: PRINCIPIO DE FUNCIONAMIENTO
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Solenoides y relés", "El solenoide: principio de funcionamiento")
add_picture_contain(s, f"{D}/solenoide_principio.png", Inches(0.6), Inches(1.9), Inches(4.6), Inches(4.6))
add_bullets(s, Inches(5.55), Inches(1.95), Inches(7.15), Inches(4.6), [
    ("Construcción.  ", "Un núcleo de hierro móvil (émbolo) dentro de una bobina de alambre, con un resorte de retorno."),
    ("Física.  ", "Al aplicar corriente, el campo magnético induce una fuerza que mueve el émbolo para minimizar la reluctancia magnética (cerrar el entrehierro)."),
    ("La fuerza es proporcional al cuadrado de la corriente:  ", "F ∝ i²."),
    ("La fuerza disminuye drásticamente ", "con la distancia del entrehierro."),
    ("Uso típico:  ", "válvulas, cerrojos y relés."),
], size=15.5, space_after=14)
add_page_footer(s, 4)
add_notes(s, "Diagrama recortado del material original del curso.")

# ====================================================================
# 5. CIRCUITO DE CONTROL PARA SOLENOIDES
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Solenoides y relés", "Circuito de control para solenoides")
add_picture_contain(s, f"{D}/circuito_control_solenoide.png", Inches(1.6), Inches(1.85), Inches(10.1), Inches(2.75))
cw = Inches(3.85)
cards = [
    ("El problema", ["El microcontrolador solo entrega ~20 mA.", "El solenoide requiere entre 500 mA y 2 A."], ORANGE_TXT),
    ("El driver", ["El transistor actúa como interruptor electrónico de alta corriente,", "controlado por la señal lógica del microcontrolador."], TEAL_TXT),
    ("La protección", ["El solenoide es un inductor (V = L di/dt).", "Al desconectar, el diodo volante disipa el pico de voltaje inverso y protege al transistor."], NAVY),
]
x0 = Inches(0.6)
for i, (title, body, color) in enumerate(cards):
    add_card(s, x0 + i * (cw + Inches(0.15)), Inches(4.75), cw, Inches(2.05), title=title, body=body,
             accent=color, fill=LIGHTBLUE, title_size=15, body_size=12.5)
add_page_footer(s, 5)
add_notes(s, "Diagrama recortado del material original del curso.")

# ====================================================================
# 6. RELÉ: FOTOS
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Solenoides y relés", "El relé: un interruptor accionado por un solenoide")
add_picture_contain(s, f"{F}/rele_partes.jpg", Inches(0.6), Inches(1.9), Inches(6.5), Inches(4.6))
add_bullets(s, Inches(7.4), Inches(1.95), Inches(5.35), Inches(4.6), [
    ("Bobina.  ", "Genera el campo magnético al energizarse (es el mismo solenoide visto en la lámina anterior)."),
    ("Armadura.  ", "Pieza móvil que el campo magnético atrae."),
    ("Resorte.  ", "Regresa la armadura a su posición de reposo al des-energizar."),
    ("Contactos.  ", "Se abren o cierran según la posición de la armadura, conmutando un circuito de potencia separado."),
    ("Uso.  ", "Permite que una señal de bajo voltaje (microcontrolador) controle cargas de alto voltaje o corriente, con aislamiento galvánico."),
], size=15, space_after=12)
add_page_footer(s, 6)
add_notes(s, "Fotografía: David Boettcher — dominio público (Wikimedia Commons, File:Relay_Parts.jpg).")

# ====================================================================
# 7. RELÉ: CONTACTOS C / NO / NC
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Solenoides y relés", "Relé: contactos común, NA y NC")
add_picture_contain(s, f"{D}/rele_contactos.png", Inches(0.6), Inches(1.95), Inches(12.1), Inches(3.3))
cw = Inches(3.85)
cards2 = [
    ("C — Común", "El punto móvil, compartido por los dos contactos fijos.", NAVY),
    ("NA — Normalmente Abierto", "Cerrado solo cuando la bobina está energizada.", ORANGE_TXT),
    ("NC — Normalmente Cerrado", "Abierto solo cuando la bobina está energizada.", TEAL_TXT),
]
x0 = Inches(0.6)
for i, (title, body, color) in enumerate(cards2):
    add_card(s, x0 + i * (cw + Inches(0.15)), Inches(5.55), cw, Inches(1.25), title=title, body=[body],
             accent=color, fill=LIGHTBLUE, title_size=14, body_size=12.5)
add_page_footer(s, 7)

# ====================================================================
# 8. DIVISORA 02 — MOTORES DC
# ====================================================================
add_divider_slide(prs, "02", "Motores DC",
                   "De la fuerza de Lorentz al control por PWM: el actuador rotativo más común en robótica.",
                   f"{F}/motor_dc_abierto.jpg",
                   credit="Foto: Christian Kral — CC BY 4.0 (Wikimedia Commons)")

# ====================================================================
# 9. CONSTRUCCIÓN INTERNA
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores DC", "Construcción interna")
add_picture_contain(s, f"{D}/motor_dc_construccion.png", Inches(0.6), Inches(1.9), Inches(12.1), Inches(3.9))
_, tf = add_textbox(s, Inches(0.6), Inches(6.0), Inches(12.1), Inches(0.8))
p = tf.paragraphs[0]; set_paragraph(p, line_spacing=1.2)
r = p.add_run(); _set_run(r, "La interacción entre el campo magnético estático del estator (imanes permanentes) y la corriente variable en el rotor (armadura) genera el torque. El conmutador y las escobillas invierten la corriente en el momento preciso para mantener el giro.",
                            size=14.5, color=NAVY)
add_page_footer(s, 9)
add_notes(s, "Diagrama recortado del material original del curso.")

# ====================================================================
# 10. FUERZA DE LORENTZ
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores DC", "La física del giro: fuerza de Lorentz")
add_picture_contain(s, f"{D}/fuerza_lorentz.png", Inches(0.6), Inches(1.9), Inches(5.4), Inches(4.6))
add_bullets(s, Inches(6.3), Inches(1.95), Inches(6.4), Inches(4.6), [
    ("Fórmula:  ", "F = I × B  (Fuerza = Corriente × Campo magnético)."),
    ("Generación de torque.  ", "Las fuerzas opuestas sobre cada lado de la espira crean un par de torsión sobre el eje."),
    ("El rol del conmutador.  ", "Para mantener el giro continuo, la dirección de la corriente debe invertirse cada 180°."),
    ("El conmutador y las escobillas ", "realizan esta inversión mecánicamente, de forma automática."),
], size=16, space_after=16)
add_page_footer(s, 10)

# ====================================================================
# 11. CONTROL DE MOTOR DC (TRANSISTOR + DIODO)
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores DC", "Control de motor DC con transistor y diodo")
add_picture_contain(s, f"{D}/control_motor_dc_circuito.png", Inches(2.0), Inches(1.9), Inches(9.3), Inches(3.7))
add_bullets(s, Inches(0.6), Inches(5.85), Inches(12.1), Inches(1.1), [
    ("PWM controla la velocidad.  ", "Un tren de pulsos en la base del transistor (Q3, NPN) enciende y apaga la corriente del motor a alta frecuencia."),
    ("El diodo D3 ", "cumple la misma función de protección que en el circuito del solenoide: disipa el pico inductivo al cortar la corriente."),
], size=14, space_after=8)
add_page_footer(s, 11)
add_notes(s, "Circuito original: © 2007 David Cook, www.RobotRoom.com (incluido en el material del curso).")

# ====================================================================
# 12. CARACTERÍSTICAS DEL MOTOR DC
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores DC", "Dos reglas de oro del motor DC")
cw = Inches(5.85)
add_card(s, Inches(0.6), Inches(2.3), cw, Inches(3.1), title="Más corriente", body=[
    "= más torque.", "", "La corriente en el rotor determina la fuerza de Lorentz y, por lo tanto, el par disponible en el eje."],
    accent=ORANGE_TXT, title_size=26, body_size=15, fill=LIGHTBLUE)
add_card(s, Inches(6.85), Inches(2.3), cw, Inches(3.1), title="Más voltaje", body=[
    "= más velocidad.", "", "El voltaje aplicado determina la velocidad de vacío del motor: a mayor voltaje, mayor velocidad angular."],
    accent=TEAL_TXT, title_size=26, body_size=15, fill=LIGHTBLUE)
add_page_footer(s, 12)

# ====================================================================
# 13. CURVA TORQUE VS. VELOCIDAD (NATIVA)
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores DC", "Curva característica: torque vs. velocidad")
add_picture_contain(s, f"{G}/torque_velocidad_dc.png", Inches(0.6), Inches(1.85), Inches(12.1), Inches(4.9))
add_page_footer(s, 13)
add_notes(s, "Gráfica generada nativamente (matplotlib) a partir de la curva lineal idealizada torque-velocidad del motor DC de imán permanente.")

# ====================================================================
# 14. INTERPRETACIÓN DE DATASHEET (TABLA NATIVA)
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores DC", "Interpretación de hoja de datos (datasheet)")
add_table(s, Inches(0.6), Inches(2.0), Inches(6.0), Inches(3.1),
           ["Parámetro", "Valor"],
           [["Voltaje nominal", "12 V"],
            ["Velocidad sin carga", "8 000 rpm"],
            ["Torque de arranque (stall)", "150 mN·m"],
            ["**Corriente de arranque (stall)**", "**5.2 A**"],
            ["Constante de torque (kt)", "25.3 mN·m/A"],
            ["Resistencia de terminales", "2.3 Ω"]],
           col_widths=[Inches(3.6), Inches(2.4)], highlight_rows=[3])
add_bullets(s, Inches(7.0), Inches(2.1), Inches(5.75), Inches(3.0), [
    ("Corriente de arranque.  ", "El driver (puente H) debe soportar esta corriente pico, no solo la nominal."),
    ("Constante de torque.  ", "Define cuánto “empuje” se obtiene por cada amperio de corriente."),
], size=14.5, space_after=14)
add_page_footer(s, 14)

# ====================================================================
# 15. DC CON ESCOBILLAS: VENTAJAS Y LIMITACIONES
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores DC", "Motor DC con escobillas: ventajas y limitaciones")
cw = Inches(5.85)
add_card(s, Inches(0.6), Inches(2.0), cw, Inches(4.5), title="Ventajas", body=[
    "Control simple: reduce velocidad al bajar el voltaje, invierte el giro al invertir la polaridad.",
    "Torque elevado a bajas velocidades.",
    "Eficiencia de aproximadamente 75-80 %.",
    "Costo muy bajo: un motor típico cuesta apenas unos pocos dólares.",
], accent=TEAL_TXT, body_size=14.5, fill=LIGHTBLUE)
add_card(s, Inches(6.85), Inches(2.0), cw, Inches(4.5), title="Limitaciones", body=[
    "Ruido audible y ruido electromagnético por las chispas en el conmutador.",
    "Ese ruido puede interferir con otras partes del sistema.",
    "Las escobillas se desgastan por el contacto continuo y requieren mantenimiento.",
    "La velocidad puede verse limitada por el calentamiento de las escobillas.",
], accent=ORANGE_TXT, body_size=14.5, fill=LIGHTBLUE)
add_page_footer(s, 15)

# ====================================================================
# 16. BLDC: PRINCIPIO
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores DC", "Motor DC sin escobillas (BLDC): principio")
add_picture_contain(s, f"{F}/motor_bldc.jpg", Inches(0.6), Inches(1.9), Inches(5.4), Inches(4.6))
add_bullets(s, Inches(6.3), Inches(1.95), Inches(6.4), Inches(4.6), [
    ("La conmutación se realiza electrónicamente ", "con controladores externos, en lugar de escobillas mecánicas."),
    ("El principio es el mismo motor de imán permanente, ", "pero invertido: los imanes están en el rotor y las bobinas en el estator."),
    ("Un controlador electrónico ", "energiza las bobinas en la secuencia correcta para mantener el giro."),
], size=16, space_after=16)
add_page_footer(s, 16)
add_notes(s, "Fotografía: Medvedev — CC BY-SA 3.0 (Wikimedia Commons, File:Stator_Winding_of_a_BLDC_Motor.jpg).")

# ====================================================================
# 17. BLDC: ROTOR / ESTATOR
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores DC", "BLDC: rotor de imanes y estator bobinado")
add_picture_contain(s, f"{D}/bldc_rotor_estator.png", Inches(1.3), Inches(1.9), Inches(10.7), Inches(4.6))
add_page_footer(s, 17)
add_notes(s, "Diagrama recortado del material original del curso. La parte difícil es controlar la polaridad de la corriente en las bobinas y mantenerla sincronizada con la velocidad del rotor.")

# ====================================================================
# 18. BLDC: CONTROL CON SENSORES HALL
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores DC", "BLDC: control con sensores de efecto Hall")
add_picture_contain(s, f"{D}/bldc_hall_cross.png", Inches(0.6), Inches(1.95), Inches(4.6), Inches(4.3))
add_picture_contain(s, f"{D}/bldc_control_bloques.png", Inches(5.5), Inches(1.95), Inches(7.2), Inches(4.3))
add_bullets(s, Inches(0.6), Inches(6.3), Inches(12.1), Inches(0.7), [
    ("Los sensores Hall ", "detectan la posición del rotor; el microcontrolador usa esa información para energizar la bobina correcta en cada instante."),
], size=13, space_after=0)
add_page_footer(s, 18)

# ====================================================================
# 19. BLDC: VENTAJAS Y LIMITACIONES
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores DC", "Motor BLDC: ventajas y limitaciones")
cw = Inches(5.85)
add_card(s, Inches(0.6), Inches(2.0), cw, Inches(4.4), title="Ventajas", body=[
    "Menos ruido eléctrico: no hay escobillas generando chispas.",
    "Más eficiente: puede sostener el par máximo de forma continua.",
    "Un motor con escobillas necesitaría un imán mucho más grande para igualar ese par.",
    "Mayor durabilidad: no hay escobillas que reemplazar.",
], accent=TEAL_TXT, body_size=14.5, fill=LIGHTBLUE)
add_card(s, Inches(6.85), Inches(2.0), cw, Inches(4.4), title="Limitaciones", body=[
    "Algunos motores BLDC son difíciles de controlar.",
    "Requieren un regulador (driver) electrónico especializado, más costoso que un simple transistor.",
], accent=ORANGE_TXT, body_size=14.5, fill=LIGHTBLUE)
add_page_footer(s, 19)

# ====================================================================
# 20. PWM (NATIVA)
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores DC", "Control de velocidad: PWM")
add_picture_contain(s, f"{G}/pwm_formas_onda.png", Inches(2.1), Inches(1.7), Inches(9.1), Inches(4.6))
_, tf = add_textbox(s, Inches(0.6), Inches(6.55), Inches(12.1), Inches(0.5))
p = tf.paragraphs[0]
r = p.add_run(); _set_run(r, "Principio: no se reduce el voltaje de la batería (ineficiente); se “corta” la energía a alta frecuencia. La inductancia del motor suaviza los pulsos en movimiento continuo.",
                            size=13, color=GRAY_TXT, italic=True)
add_page_footer(s, 20)
add_notes(s, "Gráfica generada nativamente (matplotlib).")

# ====================================================================
# 21. INTERFAZ CON MICROCONTROLADOR (BLOQUES, NATIVO)
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores DC", "Implementación práctica: interfaz con microcontrolador")


def block_diagram_micro_driver_motor(slide, y=Inches(2.4)):
    bw, bh = Inches(2.55), Inches(1.15)
    gap = Inches(0.85)
    x0 = Inches(0.9)
    labels = [
        ("Microcontrolador", "(Arduino / ESP32)", NAVY),
        ("Driver de potencia", "(L298N / DRV8825)", ORANGE_TXT),
        ("Motor DC", "", TEAL_TXT),
    ]
    xs = []
    for i, (t1, t2, color) in enumerate(labels):
        x = x0 + i * (bw + gap)
        xs.append(x)
        add_card(slide, x, y, bw, bh, title=t1, body=[t2] if t2 else None, accent=color, fill=LIGHTBLUE,
                 title_size=14.5, body_size=12, valign=MSO_ANCHOR.MIDDLE)
    # arrows between boxes
    for i in range(2):
        conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, xs[i] + bw, y + bh // 2, xs[i + 1], y + bh // 2)
        conn.line.color.rgb = NAVY
        conn.line.width = Pt(2.2)
        line = conn.line._get_or_add_ln()
        tail = line.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'})
        line.append(tail)
    # labels above arrows
    add_textbox_label(slide, xs[0] + bw + Inches(0.05), y - Inches(0.32), gap - Inches(0.1), "PWM / DIR\n(3.3V-5V lógico)", 10.5, NAVY)
    add_textbox_label(slide, xs[1] + bw + Inches(0.05), y - Inches(0.32), gap - Inches(0.1), "12V // 24V\n(potencia)", 10.5, ORANGE_TXT)

    # battery + common gnd row
    by = y + bh + Inches(0.75)
    bat = add_card(slide, xs[1], by, bw, Inches(0.95), title="Batería", body=["(fuente de poder)"], accent=NAVY,
                    fill=LIGHTBLUE, title_size=13.5, body_size=11.5, valign=MSO_ANCHOR.MIDDLE)
    conn2 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, xs[1] + bw // 2, y + bh, xs[1] + bw // 2, by)
    conn2.line.color.rgb = GRAY_TXT
    conn2.line.width = Pt(1.6)
    # common ground line across all three
    gy = by + Inches(1.15)
    gnd_line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, xs[0] + bw // 2, gy, xs[2] + bw // 2, gy)
    gnd_line.line.color.rgb = GRAY_TXT
    gnd_line.line.width = Pt(1.8)
    for x in [xs[0] + bw // 2, xs[1] + bw // 2, xs[2] + bw // 2]:
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x, y + bh, x, gy)
        c.line.color.rgb = GRAY_TXT
        c.line.width = Pt(1.4)
    add_textbox_label(slide, xs[0], gy + Inches(0.06), Inches(3.6), "Tierra común (GND) — evita diferencias de potencial entre etapas", 11.5, GRAY_TXT, align=PP_ALIGN.LEFT)


def add_textbox_label(slide, x, y, w, text, size, color, align=PP_ALIGN.CENTER):
    _, tf = add_textbox(slide, x, y, w, Inches(0.5))
    p = tf.paragraphs[0]; set_paragraph(p, align=align, line_spacing=1.05)
    lines = text.split("\n")
    for i, line in enumerate(lines):
        pp = p if i == 0 else tf.add_paragraph()
        set_paragraph(pp, align=align, line_spacing=1.05)
        r = pp.add_run(); _set_run(r, line, size=size, color=color, bold=True)


block_diagram_micro_driver_motor(s)
add_bullets(s, Inches(0.6), Inches(6.35), Inches(12.1), Inches(0.7), [
    ("Separación de potencia.  ", "La lógica (bajo voltaje) controla al driver; el driver gestiona la alta corriente de la batería hacia el motor."),
], size=13, space_after=0)
add_page_footer(s, 21)

# ====================================================================
# 22. REDUCTORAS
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores DC", "Transmisión mecánica: reductoras (gearboxes)")
add_picture_contain(s, f"{D}/reductora_gearbox.png", Inches(0.6), Inches(1.9), Inches(5.4), Inches(4.6))
add_bullets(s, Inches(6.3), Inches(1.95), Inches(6.4), Inches(4.6), [
    ("Problema.  ", "Los motores DC son de alta velocidad y bajo torque; los robots requieren baja velocidad y alto torque."),
    ("Solución.  ", "Una caja reductora con relación de reducción G."),
    ("ω_out = ω_in / G  ", "→ la velocidad disminuye."),
    ("T_out = T_in × G × η  ", "→ el torque aumenta."),
    ("η = eficiencia mecánica ", "(típicamente 60-90 %)."),
], size=15.5, space_after=13)
add_page_footer(s, 22)
add_notes(s, "Ilustración recortada del material original del curso.")

# ====================================================================
# 23. DIVISORA 03 — PUENTE H
# ====================================================================
add_divider_slide(prs, "03", "El puente H",
                   "La interfaz de potencia que traduce una señal digital en movimiento bidireccional.",
                   f"{D}/l298n_board_only.png")

# ====================================================================
# 24. EL DESAFÍO DE LA INTERFAZ DE POTENCIA
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Puente H", "El desafío de la interfaz de potencia")
add_picture_contain(s, f"{D}/arduino_blueprint.png", Inches(0.6), Inches(1.95), Inches(4.6), Inches(4.1))
add_picture_contain(s, f"{D}/motor_blueprint.png", Inches(8.1), Inches(1.95), Inches(4.6), Inches(4.1))
_, tf = add_textbox(s, Inches(5.3), Inches(2.3), Inches(2.7), Inches(3.4))
for i, (title, body, color) in enumerate([
    ("El cerebro", "Señal lógica: 5V / 3.3V\nCorriente: < 40 mA", NAVY),
    ("Incompatibilidad", "directa", ORANGE_TXT),
    ("El músculo", "Carga inductiva: 12V-48V\nCorriente: > 1A", TEAL_TXT),
]):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    set_paragraph(p, align=PP_ALIGN.CENTER, space_before=14 if i else 0, line_spacing=1.15)
    r = p.add_run(); _set_run(r, title, size=14.5, color=color, bold=True)
    for line in body.split("\n"):
        pp = tf.add_paragraph(); set_paragraph(pp, align=PP_ALIGN.CENTER, line_spacing=1.1)
        rr = pp.add_run(); _set_run(rr, line, size=11.5, color=GRAY_TXT)
add_bullets(s, Inches(0.6), Inches(6.3), Inches(12.1), Inches(0.7), [
    ("Necesitamos una etapa intermedia ", "que interprete la señal de control (información) y module la fuente de energía externa (potencia): el driver."),
], size=13.5, space_after=0)
add_page_footer(s, 24)

# ====================================================================
# 25. ARQUITECTURA DEL PUENTE H
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Puente H", "Arquitectura del puente H (H-Bridge)")
add_picture_contain(s, f"{D}/puente_h_arquitectura.png", Inches(1.4), Inches(1.85), Inches(10.5), Inches(4.4))
add_bullets(s, Inches(0.6), Inches(6.35), Inches(12.1), Inches(0.7), [
    ("Topología:  ", "4 interruptores que permiten invertir la polaridad del voltaje en la carga (el motor) usando una única fuente."),
], size=13.5, space_after=0)
add_page_footer(s, 25)

# ====================================================================
# 26. FLUJO DE CORRIENTE: CONTROL DE DIRECCIÓN
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Puente H", "Flujo de corriente: control de dirección")
add_picture_contain(s, f"{D}/puente_h_flujo_corriente.png", Inches(0.6), Inches(1.9), Inches(12.1), Inches(4.6))
add_page_footer(s, 26)
add_notes(s, "La corriente fluye en pares diagonales de transistores para invertir el sentido de giro del motor.")

# ====================================================================
# 27. MODOS DE DETENCIÓN Y ESTADOS PROHIBIDOS
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Puente H", "Modos de detención y estados prohibidos")
add_picture_contain(s, f"{D}/puente_h_modos_detencion.png", Inches(0.6), Inches(1.9), Inches(12.1), Inches(4.6))
add_page_footer(s, 27)

# ====================================================================
# 28. TABLA DE VERDAD (NATIVA)
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Puente H", "Lógica de control: tabla de verdad")
add_table(s, Inches(1.4), Inches(2.1), Inches(10.5), Inches(3.4),
           ["IN1", "IN2", "Estado del motor", "Función física"],
           [["LOW", "LOW", "Desconectado", "Coast (inercia)"],
            ["HIGH", "LOW", "Giro horario", "Avance"],
            ["LOW", "HIGH", "Giro antihorario", "Retroceso"],
            ["HIGH", "HIGH", "Cortocircuito de bobina", "Frenado"]],
           col_widths=[Inches(1.8), Inches(1.8), Inches(3.4), Inches(3.5)],
           align_cols=[PP_ALIGN.CENTER, PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT])
_, tf = add_textbox(s, Inches(1.4), Inches(5.85), Inches(10.5), Inches(0.5))
p = tf.paragraphs[0]
r = p.add_run(); _set_run(r, "Nota: IN1 e IN2 se refieren a las señales de control de entrada en el puente H del módulo L298N.",
                            size=12.5, color=GRAY_TXT, italic=True)
add_page_footer(s, 28)

# ====================================================================
# 29. MÓDULO L298N
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Puente H", "Hardware real: el módulo L298N")
add_picture_contain(s, f"{D}/l298n_modulo.png", Inches(0.6), Inches(1.9), Inches(6.7), Inches(4.6))
add_bullets(s, Inches(7.6), Inches(2.1), Inches(5.1), Inches(4.2), [
    ("Doble puente H ", "capaz de manejar hasta 35 V y 2 A por canal."),
    ("Disipador de calor ", "(tecnología Darlington): requiere disipación térmica por la caída de voltaje interna."),
    ("Pines lógicos ", "(ENA, IN1, IN2...) para controlar dirección y velocidad."),
], size=15, space_after=16)
add_page_footer(s, 29)
add_notes(s, "Ilustración recortada del material original del curso.")

# ====================================================================
# 30. CONEXIÓN ARDUINO + L298N
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Puente H", "Diagrama de conexión: Arduino + L298N")
add_picture_contain(s, f"{D}/arduino_l298n.png", Inches(0.6), Inches(1.9), Inches(12.1), Inches(4.6))
add_page_footer(s, 30)

# ====================================================================
# 31. CÓDIGO C++/ARDUINO
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Puente H", "Implementación en código (C++ / Arduino)")
add_bullets(s, Inches(0.6), Inches(2.0), Inches(4.6), Inches(4.4), [
    ("Dirección (digital).  ", "Los pines IN1 e IN2 definen el sentido de la corriente."),
    ("Velocidad (PWM).  ", "El pin ENA define el voltaje promedio entregado al motor."),
], size=15, space_after=16)
add_code_block(s, Inches(5.4), Inches(1.95), Inches(7.3), Inches(4.65), [
    "void moverMotor(int velocidad, bool sentido) {",
    "  if (sentido) {",
    "    digitalWrite(IN1, HIGH);",
    "    digitalWrite(IN2, LOW);",
    "  } else {",
    "    digitalWrite(IN1, LOW);",
    "    digitalWrite(IN2, HIGH);",
    "  }",
    "  analogWrite(ENA, velocidad); // Control de potencia",
    "}",
], title="control_puente_h.ino")
add_page_footer(s, 31)

# ====================================================================
# 32. COMPORTAMIENTO REAL: ZONA MUERTA (NATIVA)
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Puente H", "Comportamiento real: zona muerta")
add_picture_contain(s, f"{G}/zona_muerta_pwm.png", Inches(2.0), Inches(1.75), Inches(9.1), Inches(4.6))
_, tf = add_textbox(s, Inches(0.6), Inches(6.55), Inches(12.1), Inches(0.5))
p = tf.paragraphs[0]
r = p.add_run(); _set_run(r, "El motor no es lineal en el arranque: a ciclos de trabajo bajos (< 20 %) el voltaje promedio puede no superar la fricción estática.",
                            size=13, color=GRAY_TXT, italic=True)
add_page_footer(s, 32)
add_notes(s, "Gráfica generada nativamente (matplotlib).")

# ====================================================================
# 33. CONSIDERACIONES PRÁCTICAS Y PROTECCIÓN
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Puente H", "Consideraciones prácticas y protección")
cw = Inches(3.85)
imgs = [f"{D}/icono_caida_voltaje.png", f"{D}/icono_diodo_flyback.png", f"{D}/icono_corriente_max.png"]
titles = ["Caída de voltaje", "Diodos flyback", "Corriente máxima"]
bodies = [
    "El L298N pierde ~2V en calor. Una fuente de 6V entrega solo ~4V al motor.",
    "Protección esencial contra picos de voltaje inverso generados por la bobina al detenerse.",
    "Nunca exceder 2A continuos. Riesgo de quemar el puente.",
]
for i in range(3):
    x = Inches(0.6) + i * (cw + Inches(0.15))
    add_picture_contain(s, imgs[i], x, Inches(1.95), cw, Inches(2.5))
    add_bullets(s, x, Inches(4.55), cw, Inches(1.9), [titles[i] + ".  " + bodies[i]], size=13, space_after=0)
add_page_footer(s, 33)

# ====================================================================
# 34. PORTADA PARTE 2
# ====================================================================
s = blank_slide(prs, bg=NAVY)
_, tf = add_textbox(s, Inches(0.75), Inches(0.6), Inches(6.0), Inches(0.35))
p = tf.paragraphs[0]
r = p.add_run(); _set_run(r, "PARTE 2", size=13, color=ORANGE, bold=True, spacing=1.8)
_, tf = add_textbox(s, Inches(0.75), Inches(2.3), Inches(6.9), Inches(3.6))
p = tf.paragraphs[0]; set_paragraph(p, line_spacing=1.05)
r = p.add_run(); _set_run(r, "Motores paso a paso y", font=FONT_TITLE, size=36, color=WHITE, bold=True)
p2 = tf.add_paragraph(); set_paragraph(p2, line_spacing=1.05)
r2 = p2.add_run(); _set_run(r2, "servomotores", font=FONT_TITLE, size=36, color=ORANGE, bold=True)
p3 = tf.add_paragraph(); set_paragraph(p3, space_before=20, line_spacing=1.3)
r3 = p3.add_run(); _set_run(r3, "Principios de control, anatomía interna\ny aplicaciones de ingeniería.",
                              font=FONT_BODY, size=15.5, color=RGBColor(0xC9, 0xD3, 0xE0))
rounded_image_frame(s, f"{F}/stepper_nema17.jpg", Inches(7.9), Inches(0), SLIDE_W - Inches(7.9), SLIDE_H, cover=True)
add_notes(s, "Foto: oomlout — CC BY-SA 2.0 (Wikimedia Commons, File:Nema_17_Stepper_Motor.jpg).")

# ====================================================================
# 35. ROTACIÓN CONTINUA VS. POSICIONAMIENTO
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores paso a paso", "El problema: rotación continua vs. posicionamiento")
add_picture_contain(s, f"{D}/dc_motor_continuo.png", Inches(0.6), Inches(1.95), Inches(5.7), Inches(3.5))
add_picture_contain(s, f"{D}/robot_arm_joint.png", Inches(6.85), Inches(1.95), Inches(5.7), Inches(3.5))
cw = Inches(5.85)
add_card(s, Inches(0.6), Inches(5.6), cw, Inches(1.3), title="Solución 1: lazo abierto (stepper)",
         body=["La precisión es mecánica (dientes magnéticos). Asumimos que el motor obedece."],
         accent=ORANGE_TXT, title_size=13.5, body_size=12.5, fill=LIGHTBLUE)
add_card(s, Inches(6.85), Inches(5.6), cw, Inches(1.3), title="Solución 2: lazo cerrado (servo)",
         body=["La precisión es electrónica (con sensor). Verificamos que el motor obedezca."],
         accent=TEAL_TXT, title_size=13.5, body_size=12.5, fill=LIGHTBLUE)
add_page_footer(s, 35)

# ====================================================================
# 36. STEPPER: ANATOMÍA
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores paso a paso", "Motor paso a paso: anatomía mecánica")
add_picture_contain(s, f"{D}/stepper_anatomia.png", Inches(0.6), Inches(1.85), Inches(7.1), Inches(4.7))
add_card(s, Inches(7.95), Inches(1.95), Inches(4.75), Inches(2.0), title="Definición", body=[
    "Dispositivo electromagnético que convierte pulsos digitales en rotación mecánica discreta."],
    accent=NAVY, title_size=15, body_size=13.5, fill=LIGHTBLUE)
add_card(s, Inches(7.95), Inches(4.15), Inches(4.75), Inches(2.35), title="Principio de funcionamiento", body=[
    "Al energizar las bobinas en secuencia, los dientes del rotor se alinean magnéticamente con el estator, bloqueándose en una posición fija (detent).",
    "", "Resolución típica: 1.8° por paso (200 pasos/revolución)."],
    accent=ORANGE_TXT, title_size=15, body_size=13, fill=LIGHTBLUE)
add_page_footer(s, 36)
add_notes(s, "Diagrama recortado del material original del curso.")

# ====================================================================
# 37. UNIPOLAR VS. BIPOLAR
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores paso a paso", "Topologías de cableado: unipolar vs. bipolar")
add_picture_contain(s, f"{D}/unipolar_bipolar.png", Inches(0.6), Inches(1.9), Inches(12.1), Inches(3.6))
cw = Inches(5.85)
add_card(s, Inches(0.6), Inches(5.7), cw, Inches(1.2), title="Unipolar", body=[
    "5 o 6 cables, con toma central (center tap). Control simple, pero menor torque."],
    accent=NAVY, title_size=14, body_size=12.5, fill=LIGHTBLUE)
add_card(s, Inches(6.85), Inches(5.7), cw, Inches(1.2), title="Bipolar", body=[
    "4 cables. Requiere puente H. Mayor torque y eficiencia (usa el 100 % del bobinado)."],
    accent=ORANGE_TXT, title_size=14, body_size=12.5, fill=LIGHTBLUE)
add_page_footer(s, 37)
add_notes(s, "Tip de laboratorio: medir resistencia entre pares de cables para identificar bobinas.")

# ====================================================================
# 38. SECUENCIA DE FASES
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores paso a paso", "La física del paso: secuencia de fases")
add_picture_contain(s, f"{D}/secuencia_fases.png", Inches(0.6), Inches(1.9), Inches(12.1), Inches(3.5))
cw = Inches(5.85)
add_card(s, Inches(0.6), Inches(5.6), cw, Inches(1.3), title="Torque de detención", body=[
    "Resistencia magnética natural al giro cuando el motor está desenergizado (imanes permanentes)."],
    accent=NAVY, title_size=13.5, body_size=12.5, fill=LIGHTBLUE)
add_card(s, Inches(6.85), Inches(5.6), cw, Inches(1.3), title="Holding torque", body=[
    "Torque máximo para mantener la posición estática bajo carga, con las bobinas energizadas."],
    accent=TEAL_TXT, title_size=13.5, body_size=12.5, fill=LIGHTBLUE)
add_page_footer(s, 38)

# ====================================================================
# 39. WAVE DRIVE
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores paso a paso", "Modos de operación: wave drive")
add_bullets(s, Inches(0.6), Inches(2.0), Inches(6.5), Inches(1.6), [
    "Se energiza una sola fase a la vez.",
    "Es el método menos usado en la práctica, porque es muy ineficiente y produce el menor torque disponible.",
], size=15.5, space_after=12)
add_table(s, Inches(0.6), Inches(3.8), Inches(6.0), Inches(2.9),
           ["Paso", "A", "B", "C", "D"],
           [["1", "1", "0", "0", "0"], ["2", "0", "1", "0", "0"], ["3", "0", "0", "1", "0"], ["4", "0", "0", "0", "1"]],
           col_widths=[Inches(1.2)] + [Inches(1.2)] * 4,
           align_cols=[PP_ALIGN.CENTER] * 5)
add_picture_contain(s, f"{D}/wave_drive_icono.png", Inches(7.4), Inches(2.6), Inches(4.9), Inches(3.6))
add_page_footer(s, 39)

# ====================================================================
# 40. FULL STEP
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores paso a paso", "Modos de operación: full step")
add_bullets(s, Inches(0.6), Inches(2.0), Inches(6.5), Inches(1.9), [
    "Se energizan ambas fases constantemente, generando en todas las posiciones el torque completo.",
    ("En motores unipolares ", "se opera en este modo energizando una sola fase de cada bobina."),
    ("En motores bipolares ", "se energizan ambas bobinas para lograr el full step."),
], size=14.5, space_after=11)
add_table(s, Inches(0.6), Inches(4.15), Inches(6.0), Inches(2.6),
           ["Paso", "A", "B", "C", "D"],
           [["1", "1", "1", "0", "0"], ["2", "0", "1", "1", "0"], ["3", "0", "0", "1", "1"], ["4", "1", "0", "0", "1"]],
           col_widths=[Inches(1.2)] + [Inches(1.2)] * 4,
           align_cols=[PP_ALIGN.CENTER] * 5)
add_picture_contain(s, f"{D}/full_step_icono.png", Inches(7.4), Inches(2.6), Inches(4.9), Inches(3.6))
add_page_footer(s, 40)

# ====================================================================
# 41. HALF STEP
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores paso a paso", "Modos de operación: half step")
add_bullets(s, Inches(0.6), Inches(2.0), Inches(6.5), Inches(2.0), [
    "Se energiza una bobina, luego dos, y luego una de nuevo. Alternar entre una y dos fases le da al motor mayor resolución.",
    "Un motor de 200 pasos operando en este modo tendrá 400 pasos, el doble de lo normal.",
    "Sin embargo, el torque variará dependiendo de cada paso. Algunos drivers compensan esto entregando mayor corriente en el paso de una sola fase.",
], size=13.5, space_after=9)
add_table(s, Inches(7.3), Inches(1.95), Inches(3.0), Inches(4.4),
           ["Paso", "A", "B", "C", "D"],
           [["1", "1", "0", "0", "0"], ["2", "1", "1", "0", "0"], ["3", "0", "1", "0", "0"], ["4", "0", "1", "1", "0"],
            ["5", "0", "0", "1", "0"], ["6", "0", "0", "1", "1"], ["7", "0", "0", "0", "1"], ["8", "1", "0", "0", "1"]],
           col_widths=[Inches(0.7)] + [Inches(0.575)] * 4,
           align_cols=[PP_ALIGN.CENTER] * 5, body_size=11, header_size=11)
add_picture_contain(s, f"{D}/half_step_icono.png", Inches(0.6), Inches(4.35), Inches(6.2), Inches(2.9))
add_page_footer(s, 41)

# ====================================================================
# 42. ALGORITMO DE SECUENCIA
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores paso a paso", "Implementación: algoritmo de secuencia")
add_code_block(s, Inches(0.6), Inches(1.95), Inches(6.5), Inches(4.6), [
    "int steps[4][4] = {",
    "  {1, 0, 0, 0},",
    "  {0, 1, 0, 0},",
    "  {0, 0, 1, 0},",
    "  {0, 0, 0, 1}",
    "};",
    "",
    "void loop() {",
    "  for (int i = 0; i < 4; i++) {",
    "    writePort(steps[i]);",
    "    delay(speedDelay); // control de velocidad",
    "  }",
    "}",
], title="secuencia_pasos.ino", font_size=14.5)


def flowchart(slide, x, y, w):
    steps = ["Inicio", "Leer arreglo\nde pasos [i]", "Energizar\nbobinas", "Esperar\n(delay)", "¿i < 4?", "Siguiente paso"]
    bh = Inches(0.62)
    gap = Inches(0.18)
    cx = x + w // 2
    bw = Inches(2.6)
    cur_y = y
    centers = []
    for i, label in enumerate(steps):
        shape = MSO_SHAPE.ROUNDED_RECTANGLE if i in (0,) else (MSO_SHAPE.DIAMOND if "?" in label else MSO_SHAPE.RECTANGLE)
        bx = cx - bw // 2
        by = cur_y
        height = bh if "?" not in label else Inches(0.85)
        shp = slide.shapes.add_shape(shape, bx, by, bw, height)
        shp.fill.solid(); shp.fill.fore_color.rgb = LIGHTBLUE
        shp.line.color.rgb = NAVY; shp.line.width = Pt(1.4)
        shp.shadow.inherit = False
        tf = shp.text_frame; tf.word_wrap = True
        tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
        for li, line in enumerate(label.split("\n")):
            p = tf.paragraphs[0] if li == 0 else tf.add_paragraph()
            set_paragraph(p, align=PP_ALIGN.CENTER)
            r = p.add_run(); _set_run(r, line, size=11.5, color=NAVY, bold=True)
        centers.append((cx, by, by + height))
        cur_y = by + height + gap
    for i in range(len(centers) - 1):
        x1, _, y1b = centers[i]
        x2, y2t, _ = centers[i + 1]
        conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1b, x2, y2t)
        conn.line.color.rgb = ORANGE_TXT; conn.line.width = Pt(1.8)
        line = conn.line._get_or_add_ln()
        tail = line.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'})
        line.append(tail)
    # loop-back arrow from decision (derecha) hacia "Energizar bobinas" (derecha) - "Si"
    dx, dy_top, dy_bot = centers[4]
    dmid = (dy_top + dy_bot) // 2
    ex, e_top, e_bot = centers[2]
    emid = (e_top + e_bot) // 2
    turn_x = dx + bw // 2 + Inches(0.55)
    seg1 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, dx + bw // 2, dmid, turn_x, dmid)
    seg2 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, turn_x, dmid, turn_x, emid)
    seg3 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, turn_x, emid, ex + bw // 2, emid)
    for seg in (seg1, seg2, seg3):
        seg.line.color.rgb = TEAL_TXT
        seg.line.width = Pt(1.8)
    _arrow(seg3)
    add_textbox_label(slide, dx + bw // 2 + Inches(0.05), dmid - Inches(0.34), Inches(1.0), "Sí", 11, TEAL_TXT, align=PP_ALIGN.LEFT)


flowchart(s, Inches(7.6), Inches(1.75), Inches(4.6))
add_page_footer(s, 42)
add_notes(s, "La velocidad es función de la frecuencia del bucle: a menor delay, mayor velocidad.")

# ====================================================================
# 43. STEPPER: VENTAJAS Y DESVENTAJAS
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores paso a paso", "Motor paso a paso: ventajas y desventajas")
cw = Inches(5.85)
add_card(s, Inches(0.6), Inches(2.0), cw, Inches(4.5), title="Ventajas", body=[
    "Alta precisión y fácil control.",
    "Se controla mediante pulsos.",
    "Permite controlar posición, velocidad y sentido de giro.",
    "Alto torque desde el arranque.",
    "Ofrece torque de retención y torque de movimiento.",
], accent=TEAL_TXT, body_size=14, fill=LIGHTBLUE)
add_card(s, Inches(6.85), Inches(2.0), cw, Inches(4.5), title="Desventajas", body=[
    "Puede presentar resonancia si no se controla adecuadamente.",
    "No es fácil de operar a grandes velocidades.",
    "El consumo de corriente es independiente de la carga: siempre consume la corriente máxima, por lo que tiende a calentarse.",
], accent=ORANGE_TXT, body_size=14, fill=LIGHTBLUE)
add_page_footer(s, 43)

# ====================================================================
# 44. STEPPER: VENTAJAS Y DESVENTAJAS EN DETALLE
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores paso a paso", "Ventajas y desventajas en detalle")
cw = Inches(5.85)
add_card(s, Inches(0.6), Inches(2.0), cw, Inches(4.5), title="A favor", body=[
    "Compatible con sistemas digitales; se integra fácilmente con otros componentes digitales.",
    "No requiere sensores para posicionamiento ni detección de velocidad: funciona en lazo abierto.",
    "Sin problemas de inestabilidad ni necesidad de transductores de retroalimentación.",
    "Puede mover cargas inerciales y de fricción variadas; actúa como transductor de posición.",
    "Económico, fácil de usar, se detiene ante mal funcionamiento o interferencia.",
    "Buen par a baja velocidad, alta repetibilidad y larga duración; resiste sobrecargas mecánicas.",
], accent=TEAL_TXT, body_size=12.3, fill=LIGHTBLUE)
add_card(s, Inches(6.85), Inches(2.0), cw, Inches(4.5), title="En contra", body=[
    "Baja eficiencia: consume potencia máxima sin importar la carga.",
    "Si la velocidad aumenta, el par disponible disminuye.",
    "Tiende a calentarse durante su operación a alto rendimiento.",
    "Potencia de salida baja en comparación con el tamaño del motor.",
    "Precisión reducida cuando trabaja a carga completa.",
], accent=ORANGE_TXT, body_size=12.3, fill=LIGHTBLUE)
add_page_footer(s, 44)

# ====================================================================
# 45. LIMITACIÓN CRÍTICA: CURVA TORQUE-VELOCIDAD (NATIVA)
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Motores paso a paso", "Limitación crítica: curva torque-velocidad")
add_picture_contain(s, f"{G}/stepper_torque_velocidad.png", Inches(0.6), Inches(1.85), Inches(12.1), Inches(4.9))
add_page_footer(s, 45)
add_notes(s, "Gráfica generada nativamente (matplotlib). Alerta: el torque cae drásticamente a alta velocidad.")

# ====================================================================
# 46. DIVISORA 05 — SERVOMOTORES
# ====================================================================
add_divider_slide(prs, "05", "Servomotores",
                   "Cuando la posición importa: control de lazo cerrado con retroalimentación.",
                   f"{F}/servo_sg90.jpg",
                   credit="Foto: Suyash Dwivedi — CC BY-SA 4.0 (Wikimedia Commons)")

# ====================================================================
# 47. ANATOMÍA DE LAZO CERRADO
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Servomotores", "Anatomía de lazo cerrado")
add_picture_contain(s, f"{D}/servo_hobby_cutaway.png", Inches(0.6), Inches(1.9), Inches(4.6), Inches(4.6))
add_picture_contain(s, f"{D}/servo_industrial_maxon.png", Inches(8.1), Inches(1.9), Inches(4.6), Inches(4.6))
_, tf = add_textbox(s, Inches(5.35), Inches(2.3), Inches(2.65), Inches(3.6))
title_txt = "La “santísima trinidad” de componentes de un servo"
p = tf.paragraphs[0]; set_paragraph(p, align=PP_ALIGN.CENTER, line_spacing=1.1)
r = p.add_run(); _set_run(r, title_txt, size=13.5, color=NAVY, bold=True)
for i, (t, c) in enumerate([("1. Actuador (motor DC)", ORANGE_TXT), ("2. Transmisión (caja de engranajes)", TEAL_TXT), ("3. Feedback (potenciómetro / encoder)", NAVY)]):
    pp = tf.add_paragraph(); set_paragraph(pp, align=PP_ALIGN.CENTER, space_before=16)
    rr = pp.add_run(); _set_run(rr, t, size=13, color=c, bold=True)
add_page_footer(s, 47)
add_notes(s, "Ilustraciones recortadas del material original del curso (servo hobby y servo industrial Maxon).")

# ====================================================================
# 48. LAZO DE FEEDBACK (NATIVO)
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Servomotores", "Teoría de control: el lazo de feedback")


def control_loop_diagram(slide, y=Inches(2.3)):
    bw, bh = Inches(2.3), Inches(1.05)
    gap = Inches(0.55)
    x0 = Inches(1.0)
    boxes = [("Comparador", "Error e = r - y", NAVY), ("Controlador (PID)", "", ORANGE_TXT), ("Planta (motor + driver)", "", TEAL_TXT)]
    xs = []
    for i, (t, sub, color) in enumerate(boxes):
        x = x0 + i * (bw + gap) + Inches(0.9)
        xs.append(x)
        add_card(slide, x, y, bw, bh, title=t, body=[sub] if sub else None, accent=color, fill=LIGHTBLUE,
                 title_size=13, body_size=10.5, valign=MSO_ANCHOR.MIDDLE)
    # set point -> comparator
    add_textbox_label(slide, x0 - Inches(0.15), y + bh // 2 - Inches(0.42), Inches(1.0), "Set point\n(r)", 11.5, NAVY)
    c1 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x0 + Inches(0.75), y + bh // 2, xs[0], y + bh // 2)
    c1.line.color.rgb = NAVY; c1.line.width = Pt(2)
    _arrow(c1)
    for i in range(2):
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, xs[i] + bw, y + bh // 2, xs[i + 1], y + bh // 2)
        c.line.color.rgb = NAVY; c.line.width = Pt(2)
        _arrow(c)
    # output y
    out_x = xs[2] + bw + Inches(0.8)
    add_textbox_label(slide, xs[2] + bw + Inches(0.05), y + bh // 2 - Inches(0.42), Inches(1.0), "Variable de\nproceso (y)", 11.5, TEAL_TXT, align=PP_ALIGN.LEFT)
    c4 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, xs[2] + bw, y + bh // 2, out_x, y + bh // 2)
    c4.line.color.rgb = NAVY; c4.line.width = Pt(2)
    _arrow(c4)
    # feedback loop: sensor box below, from output back to comparator
    sy = y + Inches(1.9)
    sensor = add_card(slide, xs[0] + Inches(0.7), sy, Inches(2.3), Inches(0.85), title="Sensor", accent=GRAY_TXT,
                       fill=LIGHTBLUE, title_size=13, valign=MSO_ANCHOR.MIDDLE)
    # vertical line down from output point
    down_x = xs[2] + bw // 2 + Inches(0.4)
    c5 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, xs[2] + bw // 2, y + bh, xs[2] + bw // 2, sy + Inches(0.42))
    c5.line.color.rgb = GRAY_TXT; c5.line.width = Pt(1.8)
    c6 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, xs[2] + bw // 2, sy + Inches(0.42), xs[0] + Inches(0.7) + Inches(2.3), sy + Inches(0.42))
    c6.line.color.rgb = GRAY_TXT; c6.line.width = Pt(1.8)
    c7 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, xs[0] + Inches(0.7), sy + Inches(0.42), xs[0], y + bh)
    c7.line.color.rgb = GRAY_TXT; c7.line.width = Pt(1.8)
    _arrow(c7)


control_loop_diagram(s)
add_bullets(s, Inches(0.6), Inches(5.3), Inches(12.1), Inches(1.4), [
    ("El servo amplifica el voltaje ", "proporcionalmente al error."),
    ("Si una fuerza externa mueve el brazo, ", "el error aumenta y el motor lucha por regresar a la posición (rigidez o “stiffness”)."),
], size=14.5, space_after=10)
add_page_footer(s, 48)

# ====================================================================
# 49. CIRCUITO DEL SERVO ANALÓGICO
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Servomotores", "Circuito del servo analógico")
add_picture_contain(s, f"{D}/servo_circuito_detallado.png", Inches(1.4), Inches(1.85), Inches(10.5), Inches(4.9))
add_page_footer(s, 49)
add_notes(s, "Diagrama recortado del material original del curso: detector de error, amplificador (ganancia ~25) y motor DC con retroalimentación por potenciómetro.")

# ====================================================================
# 50. PWM POR POSICIÓN (NATIVA)
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Servomotores", "La señal de control: PWM por posición")
add_picture_contain(s, f"{G}/servo_pulsos_pwm.png", Inches(0.6), Inches(1.75), Inches(8.3), Inches(4.9))
add_picture_contain(s, f"{D}/servo_iconos_angulos.png", Inches(9.1), Inches(1.9), Inches(3.6), Inches(4.6))
_, tf = add_textbox(s, Inches(0.6), Inches(6.55), Inches(12.1), Inches(0.5))
p = tf.paragraphs[0]
r = p.add_run(); _set_run(r, "El ancho del pulso codifica la posición angular, no la potencia. El motor solo lee el estado ALTO.",
                            size=13, color=GRAY_TXT, italic=True)
add_page_footer(s, 50)
add_notes(s, "Gráfica generada nativamente (matplotlib).")

# ====================================================================
# 51. HOBBY (RC) VS. INDUSTRIAL
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Servomotores", "Clasificación: hobby (RC) vs. industrial")
add_picture_contain(s, f"{F}/servo_sg90.jpg", Inches(0.9), Inches(2.15), Inches(4.7), Inches(3.1))
add_picture_contain(s, f"{D}/servo_industrial_maxon.png", Inches(7.7), Inches(2.15), Inches(4.7), Inches(3.1))
cw = Inches(5.85)
add_card(s, Inches(0.6), Inches(5.4), cw, Inches(1.5), title="Hobby (RC)", body=[
    "Feedback: potenciómetro (desgaste mecánico). Rango: 0-180°.",
    "Interfaz: 3 cables (VCC, GND, señal). Uso: prototipado, grippers ligeros.",
], accent=ORANGE_TXT, title_size=14.5, body_size=12, fill=LIGHTBLUE)
add_card(s, Inches(6.85), Inches(5.4), cw, Inches(1.5), title="Industrial", body=[
    "Feedback: encoder óptico/magnético (alta precisión). Rango: rotación continua multiturno.",
    "Interfaz: bus industrial (CAN, EtherCAT). Uso: brazos robóticos 6-DOF, CNC.",
], accent=TEAL_TXT, title_size=14.5, body_size=12, fill=LIGHTBLUE)
add_page_footer(s, 51)
add_notes(s, "Foto SG90: Suyash Dwivedi — CC BY-SA 4.0 (Wikimedia Commons). Servo industrial: ilustración del material original del curso.")

# ====================================================================
# 52. CÓDIGO DE MAPEO
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Servomotores", "Código de ejemplo: mapeo y control")
add_code_block(s, Inches(1.2), Inches(1.95), Inches(10.9), Inches(3.7), [
    "int angle = 90; // Objetivo en grados",
    "",
    "// Mapear angulo a microsegundos (PWM)",
    "// 0 grados -> 1000us, 180 grados -> 2000us",
    "int pulseWidth = map(angle, 0, 180, 1000, 2000);",
    "",
    "// Generar pulso manual",
    "digitalWrite(servoPin, HIGH);",
    "delayMicroseconds(pulseWidth);",
    "digitalWrite(servoPin, LOW);",
    "delay(20 - (pulseWidth / 1000)); // completar ciclo de 20ms",
], title="mapeo_servo.ino")
_, tf = add_textbox(s, Inches(1.2), Inches(5.85), Inches(10.9), Inches(0.6))
p = tf.paragraphs[0]
r = p.add_run(); _set_run(r, "Nota: en la práctica, librerías como <Servo.h> manejan estas interrupciones automáticamente.",
                            size=12.5, color=GRAY_TXT, italic=True)
add_page_footer(s, 52)

# ====================================================================
# 53. DIVISORA 06 — ¿CUÁL ELEGIR?
# ====================================================================
add_divider_slide(prs, "06", "¿Cuál elegir?",
                   "Stepper o servo: la decisión depende de la carga, la precisión requerida y el presupuesto.",
                   f"{F}/brazo_robotico.jpg",
                   credit="Foto: Freek Wolsink — Pexels License")

# ====================================================================
# 54. COMPARATIVA DE RENDIMIENTO (NATIVA)
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Comparativa y selección", "Comparativa de rendimiento: torque vs. velocidad")
add_picture_contain(s, f"{G}/comparativa_stepper_servo.png", Inches(0.6), Inches(1.85), Inches(12.1), Inches(4.5))
_, tf = add_textbox(s, Inches(0.6), Inches(6.45), Inches(12.1), Inches(0.5))
p = tf.paragraphs[0]
r = p.add_run(); _set_run(r, "Stepper = fuerza estática. Servo = fuerza dinámica.", size=14, color=NAVY, bold=True)
add_page_footer(s, 54)
add_notes(s, "Gráfica generada nativamente (matplotlib).")

# ====================================================================
# 55. RESUMEN TÉCNICO (TABLA NATIVA)
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Comparativa y selección", "Resumen técnico: ¿cuál elegir?")
add_table(s, Inches(0.6), Inches(1.95), Inches(12.1), Inches(4.5),
           ["Característica", "Motor paso a paso (stepper)", "Servomotor"],
           [["Control", "Lazo abierto (open loop)", "Lazo cerrado (closed loop)"],
            ["Feedback", "Ninguno (posible pérdida de pasos)", "Encoder (autocorrección)"],
            ["Torque", "Alto en el arranque, bajo en velocidad", "Constante y dinámico"],
            ["Estabilidad", "Vibración / resonancia", "Alta estabilidad (PID)"],
            ["Costo", "Bajo / simple", "Alto / complejo"]],
           col_widths=[Inches(3.1), Inches(4.5), Inches(4.5)])
add_page_footer(s, 55)

# ====================================================================
# 56. ESCENARIOS DE APLICACIÓN
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Comparativa y selección", "Escenarios de aplicación")
cw = Inches(5.85)
add_picture_contain(s, f"{F}/impresora_3d.jpg", Inches(0.6), Inches(1.95), cw, Inches(3.3))
add_card(s, Inches(0.6), Inches(5.4), cw, Inches(1.4), title="Impresora 3D / CNC", body=[
    "Elección: stepper. Cargas predecibles, necesidad de precisión geométrica y bajo costo."],
    accent=ORANGE_TXT, title_size=15, body_size=13, fill=LIGHTBLUE)
add_picture_contain(s, f"{F}/brazo_robotico.jpg", Inches(6.85), Inches(1.95), cw, Inches(3.3))
add_card(s, Inches(6.85), Inches(5.4), cw, Inches(1.4), title="Brazo robótico industrial", body=[
    "Elección: servo. Cargas variables (inercia), requiere alta velocidad, aceleración y seguridad (detección de colisión)."],
    accent=TEAL_TXT, title_size=15, body_size=13, fill=LIGHTBLUE)
add_page_footer(s, 56)
add_notes(s, "Foto impresora 3D: Jakub Zerdzicki (Pexels License). Foto brazo robótico: Freek Wolsink (Pexels License).")

# ====================================================================
# 57. SELECCIÓN DE UN MOTOR (11 CRITERIOS)
# ====================================================================
s = blank_slide(prs)
add_label_title(s, "Comparativa y selección", "Selección de un motor: 11 criterios")
criteria = [
    "Torque de arranque", "Máxima velocidad", "Rango de velocidad", "Variación torque-velocidad",
    "Reversibilidad", "Ciclo de trabajo operativo", "Potencia requerida", "Alimentación",
    "Control preciso de posición", "Control preciso de velocidad", "¿Se requiere transmisión o engranaje?",
]
cols = 3
col_w = Inches(3.95)
row_h = Inches(1.15)
for i, crit in enumerate(criteria):
    col = i % cols
    row = i // cols
    x = Inches(0.6) + col * (col_w + Inches(0.15))
    y = Inches(1.95) + row * (row_h + Inches(0.15))
    card = add_rect(s, x, y, col_w, row_h, fill=LIGHTBLUE, radius=0.08)
    _, tf = add_textbox(s, x + Inches(0.2), y, col_w - Inches(0.4), row_h, anchor=MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]; set_paragraph(p, line_spacing=1.1)
    r = p.add_run(); _set_run(r, f"{i + 1:02d}", font=FONT_TITLE, size=20, color=ORANGE_TXT, bold=True)
    p2 = tf.add_paragraph(); set_paragraph(p2, space_before=2, line_spacing=1.05)
    r2 = p2.add_run(); _set_run(r2, crit, size=13, color=NAVY, bold=True)
add_page_footer(s, 57)

# ====================================================================
# GUARDAR
# ====================================================================
out_path = os.path.join(ROOT, "Actuadores - rediseño.pptx")
prs.save(out_path)
print("Guardado en", out_path)
print("Total de diapositivas:", len(prs.slides.__iter__.__self__._sldIdLst))
