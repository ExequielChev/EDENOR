import os
import subprocess

BASE_DIR = os.path.join(os.getcwd(), 'iniciar', 'tasks.robot')


def ejecutar_script(nombre_script):
    subprocess.run(nombre_script, shell=True)

if __name__ == "__main__":
    # Scripts a ejecutar
    scripts = (f'robot {BASE_DIR}')

    # Ejecutar el primer script normalmente carga devengados
    ejecutar_script(scripts)