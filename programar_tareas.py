"""
Configura las tareas programadas en Windows Task Scheduler.
Crea tareas para ejecutar el agente de valorizacion semanal automaticamente.

Horario: Lunes y Martes, cada 1 hora de 9AM a 6PM (hora Peru).

Uso:
  python programar_tareas.py          # Crear las tareas
  python programar_tareas.py --ver    # Ver estado de las tareas
  python programar_tareas.py --borrar # Eliminar las tareas
"""
import subprocess
import sys
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BAT_FILE = os.path.join(BASE_DIR, "ejecutar.bat")

# Horarios de ejecucion: cada hora de 9AM a 6PM
# Solo se ejecutan Lunes y Martes (configurado en /d LUN,MAR)
HORAS = [
    "09:00", "10:00", "11:00", "12:00", "13:00",
    "14:00", "15:00", "16:00", "17:00", "18:00",
]

# Generar lista de tareas
TAREAS = []
for hora in HORAS:
    hora_label = hora.replace(":", "")
    TAREAS.append({
        "nombre": f"ValSemanal_{hora_label}",
        "hora": hora,
        "bat": BAT_FILE,
        "descripcion": f"Valorizacion Semanal - Revision {hora}",
    })


def crear_tareas():
    """Crea las tareas programadas en Windows Task Scheduler (Lun-Mar)."""
    print("=" * 60)
    print("PROGRAMANDO TAREAS EN WINDOWS TASK SCHEDULER")
    print("Agente: Valorizacion Semanal")
    print("Dias: Lunes y Martes")
    print(f"Horario: {HORAS[0]} a {HORAS[-1]} (cada 1 hora)")
    print("=" * 60)

    exitos = 0
    for tarea in TAREAS:
        print(f"\n  Creando: {tarea['nombre']} ({tarea['hora']})...")

        # Primero intentar eliminar si ya existe
        subprocess.run(
            ["schtasks", "/delete", "/tn", tarea["nombre"], "/f"],
            capture_output=True, text=True
        )

        # Crear la tarea: SEMANAL, dias LUN y MAR
        cmd = [
            "schtasks", "/create",
            "/tn", tarea["nombre"],
            "/tr", f'"{tarea["bat"]}"',
            "/sc", "WEEKLY",
            "/d", "LUN,MAR",
            "/st", tarea["hora"],
            "/f",
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode == 0:
            print(f"  [OK] {tarea['nombre']} -> Lun/Mar a las {tarea['hora']}")
            exitos += 1
        else:
            print(f"  [ERROR] {tarea['nombre']}: {result.stderr.strip()}")

    print(f"\n{'=' * 60}")
    print(f"Resultado: {exitos}/{len(TAREAS)} tareas creadas exitosamente.")

    if exitos == len(TAREAS):
        print(f"\nHORARIOS CONFIGURADOS (Lunes y Martes):")
        for t in TAREAS:
            print(f"  {t['hora']} -> {t['descripcion']}")
    print("=" * 60)


def ver_tareas():
    """Muestra el estado actual de las tareas programadas."""
    print("ESTADO DE TAREAS PROGRAMADAS:")
    print("-" * 60)
    for tarea in TAREAS:
        result = subprocess.run(
            ["schtasks", "/query", "/tn", tarea["nombre"], "/fo", "LIST"],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            for line in result.stdout.split("\n"):
                line = line.strip()
                if any(k in line for k in ["Nombre", "TaskName", "Estado", "Status",
                                            "Hora", "Next Run", "Last Run",
                                            "Siguiente", "Ultima"]):
                    print(f"  {line}")
            print()
        else:
            print(f"  {tarea['nombre']}: NO ENCONTRADA")
            print()


def borrar_tareas():
    """Elimina todas las tareas programadas."""
    print("ELIMINANDO TAREAS PROGRAMADAS:")
    print("-" * 60)
    for tarea in TAREAS:
        result = subprocess.run(
            ["schtasks", "/delete", "/tn", tarea["nombre"], "/f"],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            print(f"  [OK] {tarea['nombre']} eliminada.")
        else:
            print(f"  [--] {tarea['nombre']} no existia.")


if __name__ == "__main__":
    if "--ver" in sys.argv:
        ver_tareas()
    elif "--borrar" in sys.argv:
        borrar_tareas()
    else:
        crear_tareas()
