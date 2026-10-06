"""Lossless display of stored numeric values; rounding requires a declared policy."""
from decimal import Decimal, ROUND_HALF_EVEN


def decimal_value(value):
    return Decimal(str(value))


def exact_number(value):
    number=decimal_value(value)
    if not number.is_finite(): raise ValueError('non-finite display value')
    if number==0: return '0'
    text=format(number,'f')
    return text.rstrip('0').rstrip('.') if '.' in text else text


def fixed_number(value,policy):
    if policy.get('mode')!='fixed' or policy.get('rounding')!='half_even':
        raise ValueError('rounded display requires an explicit fixed/half_even policy')
    places=policy['decimal_places']
    if not isinstance(places,int) or not 0<=places<=6: raise ValueError('unsupported decimal places')
    rounded=decimal_value(value).quantize(Decimal(1).scaleb(-places),rounding=ROUND_HALF_EVEN)
    return format(rounded,f'.{places}f')
