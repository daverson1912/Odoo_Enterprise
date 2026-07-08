# La tienda "Mi Tienda de Prueba" es de la empresa "EMPRESA VEN - UDS" (ID: 5)
# Los empleados estan en "My Company" (ID: 1) - por eso no aparecen

# Solucion: crear empleados en la compania 5 o cambiar la compania de los empleados existentes

company_5 = env['res.company'].browse(5)
print("Compania objetivo:", company_5.name)

# Ver si ya hay empleados en la compania 5
emp_cia5 = env['hr.employee'].search([('company_id', '=', 5)])
print("Empleados actuales en cia 5:", [(e.id, e.name) for e in emp_cia5])

# Opcion A: Crear un empleado "Daverson" en la compania 5 con PIN 1234
daverson_user = env['res.users'].search([('login', '=', '1234@gmail.com')], limit=1)
print("Usuario Daverson ID:", daverson_user.id)

# Verificar si Daverson ya pertenece a la compania 5
print("Companias de Daverson:", [c.name for c in daverson_user.company_ids])

# Agregar la compania 5 al usuario Daverson
if company_5 not in daverson_user.company_ids:
    daverson_user.write({'company_ids': [(4, company_5.id)]})
    print("Compania 5 agregada al usuario Daverson")

# Crear empleado de Daverson en compania 5
emp_daverson_cia5 = env['hr.employee'].search([
    ('user_id', '=', daverson_user.id),
    ('company_id', '=', 5)
], limit=1)

if not emp_daverson_cia5:
    new_emp = env['hr.employee'].with_company(company_5).create({
        'name': daverson_user.name,
        'user_id': daverson_user.id,
        'pin': '1234',
        'company_id': 5,
    })
    print("Empleado creado en cia 5: ID", new_emp.id, "| PIN:", new_emp.pin)
else:
    emp_daverson_cia5.write({'pin': '1234'})
    print("Empleado ya existia en cia 5, PIN actualizado:", emp_daverson_cia5.name)

# Tambien crear Mitchell Admin en compania 5
mitchell_user = env['res.users'].search([('login', '=', 'admin')], limit=1)
if company_5 not in mitchell_user.company_ids:
    mitchell_user.write({'company_ids': [(4, company_5.id)]})
    print("Compania 5 agregada al usuario Mitchell Admin")

emp_mitchell_cia5 = env['hr.employee'].search([
    ('user_id', '=', mitchell_user.id),
    ('company_id', '=', 5)
], limit=1)

if not emp_mitchell_cia5:
    new_mitchell = env['hr.employee'].with_company(company_5).create({
        'name': 'Mitchell Admin',
        'user_id': mitchell_user.id,
        'pin': '1234',
        'company_id': 5,
    })
    print("Mitchell Admin creado en cia 5: ID", new_mitchell.id)
else:
    emp_mitchell_cia5.write({'pin': '1234'})
    print("Mitchell Admin ya existia en cia 5, PIN actualizado")

env.cr.commit()
print("Cambios guardados.")

# Verificar el resultado
print("\n=== VERIFICACION FINAL ===")
from odoo.osv.expression import AND
pc = env['pos.config'].search([('name', 'like', 'Prueba')], limit=1)
domain = pc._employee_domain(pc.current_user_id.id)
employees = env['hr.employee'].search(domain)
print("Tienda:", pc.name)
print("Empleados disponibles ahora:", [(e.id, e.name, e.pin) for e in employees])
