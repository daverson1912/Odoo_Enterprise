import sys

def print_fields(model_name):
    try:
        model = env[model_name]
        print(f"FIELDS FOR {model_name}:", list(model._fields.keys()))
    except Exception as e:
        print(f"ERROR FOR {model_name}:", str(e))

print_fields('pos.payment.method')
print_fields('account.journal')
print_fields('product.category')
print_fields('product.product')
