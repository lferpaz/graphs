integrantes = ["Jordi", "Luis", "Frances"]

# Función para rotar los integrantes en las combinaciones deseadas
def rotar_integrantes(integrantes):
    turnos = []

    # Rotar dos en la mañana y uno en la tarde
    for i in range(len(integrantes)):
        manana = integrantes[i:i + 2]
        tarde = [integrantes[(i + 2) % len(integrantes)]]
        turnos.append((manana, tarde))

    return turnos

# Obtener las combinaciones de turnos
combinaciones_turnos = rotar_integrantes(integrantes)

# Mostrar las combinaciones de turnos
for i, turno in enumerate(combinaciones_turnos, start=1):
    manana, tarde = turno
    print(f"Combinación {i}: Mañana -> {', '.join(manana)}, Tarde -> {', '.join(tarde)}")
