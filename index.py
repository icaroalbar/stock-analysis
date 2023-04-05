import yfinance as yf
from exchange import get_exchange_name
from symbols import symbols
from datetime import datetime, timedelta
import os
import sys


class NullWriter:
    def write(self, s):
        pass


def get_stock_info(ticker_symbol):
    stock = yf.Ticker(ticker_symbol)

    try:
        # Obtendo informações da empresa
        info = stock.info

        # Exibindo informações solicitadas
        print("Informações da Empresa:")
        print("Símbolo:", info['symbol'])
        print("Nome:", info['longName'])
        print("Código Exchange:", info['exchange'])
        print("Exchange:", get_exchange_name(info['exchange']))

        # Calculando a data de um ano atrás
        today = datetime.now()

        # Obtendo dados históricos das ações do mesmo dia nos últimos dez anos
        print("\nDados históricos das ações do mesmo dia nos últimos dez anos:")
        for i in range(1, 11):
            n_years_ago = today - timedelta(days=365 * i)

            # Redirecionando temporariamente a saída padrão para suprimir a mensagem indesejada
            saved_stdout = sys.stdout
            sys.stdout = NullWriter()

            historical_data = stock.history(start=n_years_ago.strftime(
                '%Y-%m-%d'), end=(n_years_ago + timedelta(days=1)).strftime('%Y-%m-%d'))

            # Restaurando a saída padrão
            sys.stdout = saved_stdout

            if len(historical_data) == 0:
                print(
                    f"{n_years_ago.year}: Não há informações disponíveis para esta data.")
                continue

            open_price = historical_data.iloc[0]['Open']
            close_price = historical_data.iloc[0]['Close']
            percentage_change = ((close_price - open_price) / open_price) * 100

            print(f"{n_years_ago.year}: Variação percentual entre o preço de abertura e o preço de fechamento: {percentage_change:.2f}%")

    except Exception as e:
        print("Ocorreu um erro ao buscar informações da ação:", e)


if __name__ == "__main__":
    for ticker in symbols:
        print("\n" + "=" * 50)
        get_stock_info(ticker)
