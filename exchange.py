# exchange.py

def get_exchange_name(exchange_code):
    exchange_map = {
        'NMS': 'Nasdaq Stock Market',
        'NYQ': 'New York Stock Exchange',
        'SAO': 'São Paulo Stock Exchange',
        'LSE': 'London Stock Exchange',
        'TSE': 'Tokyo Stock Exchange',
        # Adicione outros códigos de exchange e seus nomes completos aqui, conforme necessário
    }

    return exchange_map.get(exchange_code, 'Desconhecido')
