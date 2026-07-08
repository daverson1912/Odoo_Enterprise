users = env['res.users'].search([('login', '=', 'prueba')])
if not users:
    # Si no lo encuentra por login, buscar por nombre
    users = env['res.users'].search([('name', 'ilike', 'prueba')])

if not users:
    print("NO SE ENCONTRÓ NINGÚN USUARIO 'prueba'")

for u in users:
    try:
        u.unlink()
        print("USUARIO ELIMINADO CON ÉXITO: ", u.name, " (Login: ", u.login, ")")
    except Exception as e:
        print("ERROR AL ELIMINAR EL USUARIO: ", u.name, str(e))
env.cr.commit()
