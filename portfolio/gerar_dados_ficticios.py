"""Gera uma base de vendas fictícia no formato de exportação de pedidos da VTEX.

A base imita a estrutura e os problemas de qualidade de uma base real
(pedidos divididos em pacotes, status e UFs escritos de formas diferentes,
colunas vazias), mas todos os valores são inventados.

Uso:
    python gerar_dados_ficticios.py
"""

import numpy as np
import pandas as pd

SEMENTE = 42
ARQUIVO_SAIDA = "vendas_ficticias.xlsx"
PEDIDOS_POR_DIA = 90
INICIO = "2025-01-01"
FIM = "2025-12-31"

rng = np.random.default_rng(SEMENTE)

TAMANHOS_ROUPA = ["PP", "P", "M", "G", "GG"]
TAMANHOS_CALCADO = ["37", "38", "39", "40", "41", "42", "43"]

# (modelo, preço, tamanhos, popularidade)
CATALOGO = [
    ("Tênis Corrida Leve Azul", 499.9, TAMANHOS_CALCADO, 6),
    ("Tênis Corrida Leve Preto", 499.9, TAMANHOS_CALCADO, 8),
    ("Tênis Casual Couro Branco", 399.9, TAMANHOS_CALCADO, 7),
    ("Tênis Trilha Impermeável Verde", 649.9, TAMANHOS_CALCADO, 3),
    ("Chinelo Slide Conforto Preto", 89.9, ["37/38", "39/40", "41/42", "43/44"], 14),
    ("Chinelo Slide Conforto Areia", 89.9, ["37/38", "39/40", "41/42", "43/44"], 9),
    ("Camiseta Dry Fit Masculina Preta", 129.9, TAMANHOS_ROUPA, 16),
    ("Camiseta Dry Fit Masculina Cinza", 129.9, TAMANHOS_ROUPA, 11),
    ("Camiseta Dry Fit Feminina Rosa", 129.9, TAMANHOS_ROUPA, 12),
    ("Camiseta Algodão Básica Branca", 79.9, TAMANHOS_ROUPA, 18),
    ("Regata Treino Feminina Preta", 99.9, TAMANHOS_ROUPA, 10),
    ("Legging Compressão Preta", 219.9, TAMANHOS_ROUPA, 13),
    ("Legging Compressão Vinho", 219.9, TAMANHOS_ROUPA, 6),
    ("Shorts Treino Masculino Preto", 149.9, TAMANHOS_ROUPA, 11),
    ("Shorts Corrida Feminino Azul", 139.9, TAMANHOS_ROUPA, 8),
    ("Top Esportivo Alta Sustentação", 179.9, TAMANHOS_ROUPA, 9),
    ("Jaqueta Corta-Vento Azul Marinho", 449.9, TAMANHOS_ROUPA, 5),
    ("Jaqueta Puffer Feminina Preta", 799.9, TAMANHOS_ROUPA, 4),
    ("Moletom Canguru Cinza Mescla", 299.9, TAMANHOS_ROUPA, 7),
    ("Calça Jogger Masculina Preta", 259.9, TAMANHOS_ROUPA, 7),
    ("Macacão Fitness Feminino Preto", 289.9, TAMANHOS_ROUPA, 5),
    ("Meia Cano Médio Kit 3 Pares", 59.9, ["P", "M", "G"], 15),
    ("Boné Aba Curva Preto", 119.9, ["UNICO"], 9),
    ("Viseira Corrida Branca", 99.9, ["UNICO"], 5),
    ("Garrafa Térmica 750ml Aço", 149.9, ["UNICO"], 12),
    ("Copo Térmico 1,2L Areia", 169.9, ["UNICO"], 10),
    ("Nécessaire Viagem Preta", 129.9, ["UNICO"], 13),
    ("Pochete Esportiva Preta", 139.9, ["UNICO"], 8),
    ("Mochila Urbana 25L Preta", 399.9, ["UNICO"], 7),
    ("Mala de Bordo Rígida Grafite", 1299.9, ["UNICO"], 4),
    ("Mala Duffel Acolchoada Preta", 899.9, ["UNICO"], 6),
    ("Mala Duffel Acolchoada Azul", 899.9, ["UNICO"], 4),
    ("Bolsa Tote Nylon Preta", 349.9, ["UNICO"], 6),
    ("Óculos de Sol Esportivo", 459.9, ["UNICO"], 3),
    ("Relógio Monitor Cardíaco", 1499.9, ["UNICO"], 2),
]

# (UF, região, peso)
ESTADOS = [
    ("SP", "Sudeste", 40), ("MG", "Sudeste", 12), ("RJ", "Sudeste", 11), ("ES", "Sudeste", 1.5),
    ("PR", "Sul", 5), ("SC", "Sul", 3.5), ("RS", "Sul", 4),
    ("BA", "Nordeste", 5), ("PE", "Nordeste", 2.5), ("CE", "Nordeste", 2), ("PB", "Nordeste", 1),
    ("RN", "Nordeste", 0.9), ("AL", "Nordeste", 0.6), ("SE", "Nordeste", 0.6), ("MA", "Nordeste", 0.7),
    ("PI", "Nordeste", 0.5), ("DF", "Centro Oeste", 4), ("GO", "Centro Oeste", 2),
    ("MT", "Centro Oeste", 1), ("MS", "Centro Oeste", 1), ("PA", "Norte", 1), ("AM", "Norte", 0.9),
    ("RO", "Norte", 0.3), ("TO", "Norte", 0.3), ("AC", "Norte", 0.1), ("AP", "Norte", 0.1), ("RR", "Norte", 0.1),
]

# Grafias "sujas" que aparecem na UF, como numa base real preenchida à mão
UF_SUJA = {
    "SP": ["sp", "Sp", "SP ", "São Paulo", "SÃO PAULO "],
    "MG": ["mg", "Mg ", "Minas Gerais"],
    "RJ": ["rj", "Rio de Janeiro"],
    "BA": ["bahia", "ba "],
    "PE": ["pe", "Pé"],
    "PA": ["Pará", "pa"],
}

LOJAS = [
    "Loja Online", "CD São Paulo", "Loja Shopping Norte", "Loja Centro BH",
    "Loja Barra RJ", "Loja Curitiba", "Loja Salvador", "Loja Brasília",
]

CANAIS = ["Site", "App", "Marketplace", "Lojas (Ship from Store)"]
PAGAMENTOS = ["Pix", "Cartão de Crédito", "Boleto"]
MOTIVOS_CANCELAMENTO = ["Atraso na entrega", "Desistência do cliente", "Pagamento não aprovado", "Falta de estoque"]

# Datas de campanha e o multiplicador de pedidos do dia
CAMPANHAS = {
    "2025-03-15": 3.0, "2025-03-16": 1.8,                      # Dia do Consumidor
    "2025-05-08": 2.0, "2025-05-09": 2.2, "2025-05-10": 1.8,   # Dia das Mães
    "2025-11-28": 6.0, "2025-11-29": 3.5, "2025-11-30": 2.5,   # Black Friday
    "2025-12-01": 2.2,                                         # Cyber Monday
}

# Taxa de cancelamento por mês: sobe no meio do ano (problema logístico) e volta a cair
CANCELAMENTO_MES = {1: 0.03, 2: 0.03, 3: 0.035, 4: 0.04, 5: 0.05, 6: 0.09,
                    7: 0.16, 8: 0.14, 9: 0.08, 10: 0.05, 11: 0.06, 12: 0.045}

STATUS_ATIVOS = {
    "Faturado": 0.86, "Pagamento Aprovado": 0.05, "Pagamento aprovado": 0.01,
    "Preparando Entrega": 0.04, "Preparando entrega": 0.01, "Pagamento Pendente": 0.005,
    "Verificando Fatura": 0.004, "Pronto para o manuseio": 0.004, "Pronto para manuseio": 0.002,
    "Aguardando decisão do Seller": 0.002, "Aguardando troca do seller": 0.002,
    "Verificando nota fiscal": 0.002,
}
STATUS_CANCELADOS = {
    "Cancelado": 0.97, "Cancelar": 0.012, "Cancelamento Solicitado": 0.008,
    "Aguardando para alterar seller": 0.005, "admin / orders.order-status.": 0.005,
}


def sortear(opcoes, pesos, n=None):
    pesos = np.asarray(pesos, dtype=float)
    return rng.choice(opcoes, size=n, p=pesos / pesos.sum())


def pedidos_por_dia():
    dias = pd.date_range(INICIO, FIM, freq="D")
    tendencia = np.linspace(1.0, 1.35, len(dias))
    fim_de_semana = np.where(dias.dayofweek >= 5, 0.85, 1.0)
    campanha = np.array([CAMPANHAS.get(d.strftime("%Y-%m-%d"), 1.0) for d in dias])
    media = PEDIDOS_POR_DIA * tendencia * fim_de_semana * campanha
    return dias, rng.poisson(media)


def gerar():
    dias, qtd_por_dia = pedidos_por_dia()
    indices_produto = np.arange(len(CATALOGO))
    popularidade = [p[3] for p in CATALOGO]
    ufs = [e[0] for e in ESTADOS]
    peso_uf = [e[2] for e in ESTADOS]
    regiao_uf = {e[0]: e[1] for e in ESTADOS}

    linhas = []
    sequencia = 5_000_000

    for dia, qtd in zip(dias, qtd_por_dia):
        for _ in range(qtd):
            sequencia += 1
            numero_pedido = f"25{sequencia:09d}"
            uf = sortear(ufs, peso_uf)
            uf_texto = uf
            if uf in UF_SUJA and rng.random() < 0.03:
                uf_texto = rng.choice(UF_SUJA[uf])
            canal = sortear(CANAIS, [45, 25, 15, 15])
            pagamento = sortear(PAGAMENTOS, [45, 48, 7])
            cupom = rng.choice(["BEMVINDO10", "FRETEGRATIS", "VOLTA15"]) if rng.random() < 0.15 else None
            qtd_pacotes = sortear([1, 2, 3], [86, 11, 3])

            for numero_pacote in range(1, qtd_pacotes + 1):
                cancelado = rng.random() < CANCELAMENTO_MES[dia.month]
                if cancelado:
                    status = sortear(list(STATUS_CANCELADOS), list(STATUS_CANCELADOS.values()))
                    motivo = sortear(MOTIVOS_CANCELAMENTO, [6, 2, 1, 1] if dia.month in (6, 7, 8) else [1, 2, 1, 1])
                else:
                    status = sortear(list(STATUS_ATIVOS), list(STATUS_ATIVOS.values()))
                    motivo = None

                loja = LOJAS[0] if numero_pacote == 1 and rng.random() < 0.6 else sortear(LOJAS, [1] * len(LOJAS))
                frete = 0.0 if rng.random() < 0.6 else float(rng.choice([9.9, 14.9, 19.9, 29.9, 39.9]))
                qtd_itens = sortear([1, 2, 3, 4], [62, 25, 9, 4])

                itens = []
                for i in sortear(indices_produto, popularidade, qtd_itens):
                    modelo, preco, tamanhos, _ = CATALOGO[i]
                    desconto = rng.choice([0, 0, 0, 0.1, 0.15, 0.2, 0.3])
                    preco_venda = round(preco * (1 - desconto), 2)
                    pecas = int(sortear([1, 2, 3], [88, 9, 3]))
                    itens.append((f"{modelo} - {rng.choice(tamanhos)}", preco, preco_venda, pecas))

                valor_pacote = round(sum(p * q for _, _, p, q in itens) + frete, 2)

                for nome, preco, preco_venda, pecas in itens:
                    linhas.append({
                        "Origin": "Marketplace",
                        "Order": f"{numero_pedido}-{numero_pacote:02d}",
                        "Sequence": sequencia * 10 + numero_pacote,
                        "Creation D": dia,
                        "UF": uf_texto,
                        "Status": status,
                        "Payment System Name": pagamento,
                        "Coupon": cupom,
                        "Peças": pecas,
                        "SKU Name": nome,
                        "SKU Value": preco,
                        "SKU Selling Price": preco_venda,
                        "SKU Total Price": round(preco_venda * pecas, 2),
                        "Shipping Value": frete,
                        "Total Value": valor_pacote,
                        "Seller Name": loja,
                        "Cancellation Reason": motivo,
                        "Canal": canal,
                        "Região": regiao_uf[uf],
                        "GiftRegistry ID": None,
                        "List Id": None,
                        "Taxes": None,
                    })

    return pd.DataFrame(linhas)


if __name__ == "__main__":
    base = gerar()
    base.to_excel(ARQUIVO_SAIDA, sheet_name="BD", index=False)
    print(f"{ARQUIVO_SAIDA}: {len(base):,} linhas, {base['Order'].nunique():,} pacotes")
