# coding=utf-8
#
# Author: Sébastien de Menten <sdementen@gmail.com>
#

# list to which append tuple (Class, function) that will be monkey patech as Class.function
monkey_patch_register = []

def enable_monkey_patch():
    """Enable monkey patching for functions declared in the ext module and registred in the monkey_patch_register list
    """
    for obj, fct in monkey_patch_register:
        setattr(obj, fct.__name__, fct)
