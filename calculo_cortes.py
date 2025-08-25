# calculo_cortes.py

import rectpack

def calcular_plano_de_corte(chapa_largura, chapa_altura, pecas):
    """
    Calcula o plano de corte, atribuindo um ID numérico único para cada peça.
    """
    
    # --- MUDANÇA NA GERAÇÃO DE ID ---
    peca_unica_id = 1 # Inicializa o contador único
    retangulos_para_alocar = []
    id_peca_map = {}
    for i, peca in enumerate(pecas):
        for j in range(peca['quantidade']):
            # Atribui o ID único e converte para string
            rid = str(peca_unica_id)
            
            retangulos_para_alocar.append((peca['largura'], peca['altura'], rid))
            id_peca_map[rid] = {'largura': peca['largura'], 'altura': peca['altura']}
            
            # Incrementa o contador para a próxima peça
            peca_unica_id += 1
    # --- FIM DA MUDANÇA ---

    packer = rectpack.newPacker(rotation=True)

    for r in retangulos_para_alocar:
        packer.add_rect(r[0], r[1], rid=r[2])

    total_de_retangulos = len(retangulos_para_alocar)
    
    while True:
        packer.add_bin(chapa_largura, chapa_altura)
        packer.pack()
        retangulos_alocados_count = sum(len(chapa) for chapa in packer)
        if retangulos_alocados_count == total_de_retangulos:
            break

    planos_agrupados = {}
    area_total_utilizada = 0
    area_chapa = chapa_largura * chapa_altura

    for chapa in packer:
        pecas_na_chapa = sorted([(r.x, r.y, r.width, r.height) for r in chapa])
        assinatura = tuple(pecas_na_chapa)

        if assinatura not in planos_agrupados:
            area_chapa_utilizada = sum(r.width * r.height for r in chapa)
            plano_de_corte = []
            pecas_contagem = {}

            for retangulo in chapa:
                tipo_key = f"{id_peca_map[retangulo.rid]['largura']}x{id_peca_map[retangulo.rid]['altura']}"
                pecas_contagem[tipo_key] = pecas_contagem.get(tipo_key, 0) + 1
                
                plano_de_corte.append({
                    "x": retangulo.x, "y": retangulo.y,
                    "largura": retangulo.width, "altura": retangulo.height,
                    "rid": retangulo.rid,
                    "tipo_largura": id_peca_map[retangulo.rid]['largura'],
                    "tipo_altura": id_peca_map[retangulo.rid]['altura']
                })
            
            resumo_pecas = [{"tipo": tipo, "qtd": qtd} for tipo, qtd in pecas_contagem.items()]

            planos_agrupados[assinatura] = {
                "plano": plano_de_corte,
                "repeticoes": 1,
                "sobra_area": area_chapa - area_chapa_utilizada,
                "resumo_pecas": resumo_pecas
            }
        else:
            planos_agrupados[assinatura]["repeticoes"] += 1
        
        area_total_utilizada += sum(r.width * r.height for r in chapa)
        
    planos_unicos = list(planos_agrupados.values())

    total_chapas = len(packer)
    area_total_chapas = total_chapas * chapa_largura * chapa_altura
    aproveitamento_geral = (area_total_utilizada / area_total_chapas) * 100 if area_total_chapas > 0 else 0

    return {
        "planos_unicos": planos_unicos,
        "total_chapas": total_chapas,
        "aproveitamento_geral": f"{aproveitamento_geral:.2f}%"
    }
