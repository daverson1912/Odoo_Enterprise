targets = env['res.company'].browse([3, 4])
for t in targets:
    try:
        t.unlink()
        print("SUCCESSFULLY DELETED: ", t.name)
    except Exception as e:
        print("FAILED TO DELETE: ", t.name, str(e))
env.cr.commit()
