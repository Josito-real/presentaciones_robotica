"""Genera las gráficas técnicas nativas (matplotlib) para el rediseño."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
import os

OUT = os.path.join(os.path.dirname(__file__), "..", "imagenes", "graficas")
os.makedirs(OUT, exist_ok=True)

# Paleta
NAVY = "#14233B"
PAPER = "#F5F3EE"
LIGHTBLUE = "#E9EEF3"
ORANGE = "#E0692A"
ORANGE_TXT = "#B8501A"
TEAL = "#1B8A8C"
TEAL_TXT = "#146E70"
GRAY = "#8A94A6"

for fpath in [
    "/usr/share/fonts/truetype/custom/SpaceGrotesk-400.ttf",
    "/usr/share/fonts/truetype/custom/SpaceGrotesk-700.ttf",
    "/usr/share/fonts/truetype/custom/IBMPlexSans-400.ttf",
    "/usr/share/fonts/truetype/custom/IBMPlexSans-700.ttf",
]:
    fm.fontManager.addfont(fpath)

plt.rcParams["font.family"] = "IBM Plex Sans"
plt.rcParams["text.color"] = NAVY
plt.rcParams["axes.edgecolor"] = NAVY
plt.rcParams["axes.labelcolor"] = NAVY
plt.rcParams["xtick.color"] = NAVY
plt.rcParams["ytick.color"] = NAVY
plt.rcParams["axes.linewidth"] = 1.4

TITLE_FONT = fm.FontProperties(fname="/usr/share/fonts/truetype/custom/SpaceGrotesk-700.ttf")


def style_ax(ax, grid=True):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(NAVY)
    ax.spines["bottom"].set_color(NAVY)
    if grid:
        ax.grid(True, color=LIGHTBLUE, linewidth=1.1, zorder=0)
    ax.set_axisbelow(True)


# ---------------------------------------------------------------
# 1. Curva torque vs velocidad motor DC
# ---------------------------------------------------------------
fig, ax = plt.subplots(figsize=(9.0, 5.4), dpi=200)
fig.patch.set_facecolor("white")
ax.set_facecolor("white")

torque = np.linspace(0, 10, 200)
omega_max = 9.0
speed = omega_max * (1 - torque / 10)
ax.plot(torque, speed, color=NAVY, linewidth=3, zorder=3)
ax.fill_between(torque, speed, color=TEAL, alpha=0.12, zorder=1)

# zona de peligro (stall)
danger_start = 7.8
ax.axvspan(danger_start, 10, color=ORANGE, alpha=0.15, zorder=1)
ax.annotate(
    "Zona de peligro\n$I_{stall}$  ·  $P = I^2R$",
    xy=(9.3, omega_max * (1 - 9.3 / 10) + 0.3), xytext=(6.0, 6.3),
    fontsize=13, color=ORANGE_TXT, fontweight="bold",
    arrowprops=dict(arrowstyle="->", color=ORANGE_TXT, lw=2),
)

ax.annotate(
    "$\\omega_{max}$ (vacío)\nmáx. velocidad, 0 torque",
    xy=(0, omega_max), xytext=(0.5, omega_max + 0.35),
    fontsize=12, color=TEAL_TXT, fontweight="bold",
)
ax.scatter([0], [omega_max], color=TEAL, s=70, zorder=4)

ax.annotate(
    "Torque de arranque ($T_s$)\nmáx. torque, 0 velocidad",
    xy=(10, 0), xytext=(5.6, 1.1),
    fontsize=12, color=NAVY, fontweight="bold",
    arrowprops=dict(arrowstyle="->", color=NAVY, lw=1.6),
)
ax.scatter([10], [0], color=NAVY, s=70, zorder=4)

ax.set_xlim(0, 10.3)
ax.set_ylim(0, 10.2)
ax.set_xlabel("Torque (N·m)", fontsize=13, fontweight="bold")
ax.set_ylabel("Velocidad angular $\\omega$", fontsize=13, fontweight="bold")
ax.set_xticks([])
ax.set_yticks([])
style_ax(ax)
plt.tight_layout()
plt.savefig(f"{OUT}/torque_velocidad_dc.png", facecolor="white")
plt.close()

# ---------------------------------------------------------------
# 2. Zona muerta por fricción estática (PWM < 20%)
# ---------------------------------------------------------------
fig, ax = plt.subplots(figsize=(9.0, 5.0), dpi=200)
duty = np.linspace(0, 100, 300)
dead_zone = 20
speed = np.where(duty < dead_zone, 0, (duty - dead_zone) / (100 - dead_zone) * 100)
ax.plot(duty, speed, color=NAVY, linewidth=3, zorder=3)
ax.axvspan(0, dead_zone, color=ORANGE, alpha=0.18, zorder=1)
ax.axvline(dead_zone, color=ORANGE_TXT, linewidth=1.6, linestyle="--", zorder=2)
ax.annotate(
    "Zona muerta\n(fricción estática)\nPWM < 20 %",
    xy=(10, 5), xytext=(2, 45),
    fontsize=12.5, color=ORANGE_TXT, fontweight="bold",
    arrowprops=dict(arrowstyle="->", color=ORANGE_TXT, lw=1.8),
)
ax.annotate(
    "El motor arranca a girar\nsolo cuando el voltaje\npromedio supera la fricción",
    xy=(45, 30), xytext=(45, 8),
    fontsize=11.5, color=NAVY,
)
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.set_xlabel("Ciclo de trabajo — Duty Cycle (%)", fontsize=13, fontweight="bold")
ax.set_ylabel("Velocidad real del motor (%)", fontsize=13, fontweight="bold")
style_ax(ax)
plt.tight_layout()
plt.savefig(f"{OUT}/zona_muerta_pwm.png", facecolor="white")
plt.close()

# ---------------------------------------------------------------
# 3. Formas de onda PWM 25/50/75%
# ---------------------------------------------------------------
fig, axes = plt.subplots(3, 1, figsize=(9.0, 6.2), dpi=200, sharex=True)
t = np.linspace(0, 4, 2000)
period = 1.0
for ax, duty, label, color in zip(
    axes, [0.25, 0.5, 0.75], ["25 %   V_prom = 0.25 x Vin", "50 %   V_prom = 0.50 x Vin", "75 %   V_prom = 0.75 x Vin"], [ORANGE, TEAL, NAVY]
):
    signal = ((t % period) < (duty * period)).astype(float)
    ax.plot(t, signal, color=color, linewidth=2.6, drawstyle="steps-post")
    ax.fill_between(t, signal, color=color, alpha=0.15, step="post")
    ax.axhline(duty, color=color, linewidth=1.3, linestyle="--", alpha=0.7)
    ax.set_ylim(-0.15, 1.3)
    ax.set_yticks([0, 1])
    ax.set_yticklabels(["0", "V_in"])
    ax.set_ylabel(label, fontsize=11, fontweight="bold", color=NAVY, rotation=0, ha="right", va="center", labelpad=10)
    style_ax(ax, grid=False)
    ax.grid(True, axis="x", color=LIGHTBLUE, linewidth=1)
axes[-1].set_xlabel("Tiempo", fontsize=13, fontweight="bold")
axes[-1].set_xticks([])
plt.tight_layout()
plt.savefig(f"{OUT}/pwm_formas_onda.png", facecolor="white")
plt.close()

# ---------------------------------------------------------------
# 4. Pulsos del servo (1.0 / 1.5 / 2.0 ms, periodo 20 ms)
# ---------------------------------------------------------------
fig, axes = plt.subplots(3, 1, figsize=(9.0, 6.2), dpi=200, sharex=True)
period_ms = 20
t = np.linspace(0, 40, 4000)
pulses = [(1.0, "-90°"), (1.5, "0° (centro)"), (2.0, "+90°")]
colors = [ORANGE, TEAL, NAVY]
for ax, (pw, angle), color in zip(axes, pulses, colors):
    signal = np.zeros_like(t)
    for start in [0, period_ms]:
        signal[(t >= start) & (t < start + pw)] = 1
    ax.plot(t, signal, color=color, linewidth=2.6, drawstyle="steps-post")
    ax.fill_between(t, signal, color=color, alpha=0.15, step="post")
    ax.annotate(f"{pw:.1f} ms  =  {angle}", xy=(pw / 2, 1.12), fontsize=12, color=color, fontweight="bold", ha="center")
    ax.set_ylim(-0.15, 1.5)
    ax.set_yticks([])
    style_ax(ax, grid=False)
axes[-1].set_xlabel("Tiempo (ms)  —  periodo total: 20 ms (50 Hz)", fontsize=12.5, fontweight="bold")
axes[-1].set_xticks([0, 20, 40])
plt.tight_layout()
plt.savefig(f"{OUT}/servo_pulsos_pwm.png", facecolor="white")
plt.close()

# ---------------------------------------------------------------
# 5. Curva torque-velocidad del stepper
# ---------------------------------------------------------------
fig, ax = plt.subplots(figsize=(9.0, 5.4), dpi=200)
rpm = np.linspace(0, 10, 300)
holding = 8.5
torque_curve = holding / (1 + (rpm / 2.2) ** 2.1)
ax.plot(rpm, torque_curve, color=NAVY, linewidth=3, zorder=3)
ax.fill_between(rpm, torque_curve, color=TEAL, alpha=0.12, zorder=1)

ax.axvspan(0, 2.2, color=TEAL, alpha=0.13, zorder=1)
ax.axvspan(2.2, 5.0, color=ORANGE, alpha=0.10, zorder=1)
ax.axvspan(5.0, 10, color=GRAY, alpha=0.12, zorder=1)

ax.annotate("Holding torque\n(máx.)", xy=(0.05, holding), xytext=(0.1, holding + 0.3),
            fontsize=12, color=TEAL_TXT, fontweight="bold")
ax.annotate("Locked-step\n(arranque / parada)", xy=(1.2, 6.0), xytext=(1.1, 3.0),
            fontsize=11.5, color=ORANGE_TXT, fontweight="bold",
            arrowprops=dict(arrowstyle="->", color=ORANGE_TXT, lw=1.6))
ax.annotate("Slewing\n(requiere rampa\nde aceleración)", xy=(7, 0.7), xytext=(6.3, 2.4),
            fontsize=11.5, color=GRAY, fontweight="bold",
            arrowprops=dict(arrowstyle="->", color=GRAY, lw=1.6))

ax.set_xlim(0, 10)
ax.set_ylim(0, holding + 1)
ax.set_xlabel("Velocidad (RPM)", fontsize=13, fontweight="bold")
ax.set_ylabel("Torque (N·m)", fontsize=13, fontweight="bold")
ax.set_xticks([])
ax.set_yticks([])
style_ax(ax)
plt.tight_layout()
plt.savefig(f"{OUT}/stepper_torque_velocidad.png", facecolor="white")
plt.close()

# ---------------------------------------------------------------
# 6. Comparativa stepper vs servo
# ---------------------------------------------------------------
fig, ax = plt.subplots(figsize=(9.0, 5.4), dpi=200)
rpm = np.linspace(0.01, 10, 300)
stepper_t = 8.0 / (1 + (rpm / 2.0) ** 2.0)
servo_t = np.where(rpm < 6.5, 5.6, 5.6 * np.exp(-(rpm - 6.5) / 2.2))
ax.plot(rpm, stepper_t, color=ORANGE, linewidth=3, label="Motor paso a paso (stepper)", zorder=3)
ax.plot(rpm, servo_t, color=TEAL, linewidth=3, label="Servomotor", zorder=3)
ax.set_xlim(0, 10)
ax.set_ylim(0, 9)
ax.set_xlabel("Velocidad del motor (RPM)", fontsize=13, fontweight="bold")
ax.set_ylabel("Torque (N·m)", fontsize=13, fontweight="bold")
ax.set_xticks([])
ax.set_yticks([])
style_ax(ax)
leg = ax.legend(loc="upper right", fontsize=12.5, frameon=False)
for text, color in zip(leg.get_texts(), [ORANGE_TXT, TEAL_TXT]):
    text.set_color(color)
    text.set_fontweight("bold")
plt.tight_layout()
plt.savefig(f"{OUT}/comparativa_stepper_servo.png", facecolor="white")
plt.close()

print("Gráficas generadas en", os.path.abspath(OUT))
for f in sorted(os.listdir(OUT)):
    print(" -", f)
