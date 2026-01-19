from django import template

register = template.Library()

@register.filter
def lookup(dictionary, key):
    """Permite acessar valores de dicionário com chaves que contêm espaços"""
    return dictionary.get(key, '')
