from pathlib import Path
import re
import unicodedata
import pandas as pd

PASTA_ENTRADA = r"C:\Users\Filippo Di Giuseppe\Downloads\arquivosConverter"
PASTA_SAIDA = r"C:\Users\Filippo Di Giuseppe\Downloads\arquivosConvertido"
ABA = "COM.+PRIV.+FAC.+PESQ."
SEPARADOR_CSV = ","
ENCODING = "utf-8-sig"


def normalizar_texto(valor: str) -> str:
    valor = str(valor or "").strip().lower()
    valor = unicodedata.normalize("NFD", valor)
    valor = "".join(ch for ch in valor if unicodedata.category(ch) != "Mn")
    valor = re.sub(r"[^\w\s]+", " ", valor)
    valor = re.sub(r"\s+", " ", valor).strip()
    return valor


def nome_coluna_padrao(coluna: str) -> str:
    base = normalizar_texto(coluna).replace(" ", "_")
    mapa = {
        "idp": "idp",
        "estado": "estado",
        "cidade": "cidade",
        "cod_estado": "cod_estado",
        "cod_municipio": "cod_municipio",
        "cod_meso": "cod_meso",
        "cod_micro": "cod_micro",
        "regiao_cidade": "regiao_cidade",
        "setor_dentro_da_cidade": "setor_dentro_cidade",
        "setor_dentro_cidade": "setor_dentro_cidade",
        "sexo": "sexo",
        "idade": "idade",
        "grau_de_instrucao": "grau_instrucao",
        "grau_instrucao": "grau_instrucao",
        "ocupacao": "ocupacao",
        "renda_familiar": "renda_familiar",
        "data_da_pesquisa": "data_pesquisa",
        "categoria": "categoria",
        "nome": "nome",
        "endereco_completo": "endereco_completo",
        "endereco": "endereco",
        "numero": "numero",
        "bairro": "bairro",
        "cep": "cep",
        "email": "email",
        "tf1": "tf1",
        "tf2": "tf2",
        "tf3": "tf3",
        "tf4": "tf4",
        "n_pesq": "numero_pesq",
        "dia_pesq": "dia_pesq",
        "qt_tf1": "qt_tf1",
        "qt_tf2": "qt_tf2",
        "qt_tf3": "qt_tf3",
        "qt_tf4": "qt_tf4",
    }
    return mapa.get(base, base)


def limpar_telefone(valor) -> str:
    if pd.isna(valor):
        return ""
    s = str(valor).strip()
    if s.lower() in {"nan", "none", ""}:
        return ""
    return re.sub(r"\D", "", s)


def linha_tem_conteudo(linha) -> bool:
    for valor in linha:
        if pd.isna(valor):
            continue
        if str(valor).strip() != "":
            return True
    return False


def processar_arquivo(caminho_xlsx: Path, pasta_saida: Path):
    print(f"Processando: {caminho_xlsx.name}")

    df = pd.read_excel(
        caminho_xlsx,
        sheet_name=ABA,
        header=1,
        dtype=object
    )

    # remove colunas totalmente vazias
    df = df.dropna(axis=1, how="all")

    # remove colunas sem nome útil
    df = df.loc[:, ~df.columns.isna()]

    # padroniza nomes
    df.columns = [nome_coluna_padrao(c) for c in df.columns]

    # adiciona origem
    df["arquivo_origem"] = caminho_xlsx.name

    # limpa telefones
    for col in ["tf1", "tf2", "tf3", "tf4"]:
        if col in df.columns:
            df[col] = df[col].apply(limpar_telefone)

    # remove linhas totalmente vazias
    df = df.dropna(how="all")
    df = df[df.apply(linha_tem_conteudo, axis=1)]

    # salva csv
    caminho_saida = pasta_saida / caminho_xlsx.with_suffix(".csv").name
    df.to_csv(caminho_saida, index=False, sep=SEPARADOR_CSV, encoding=ENCODING)

    print(f"  -> CSV salvo em: {caminho_saida}")
    print(f"  -> Linhas: {len(df)}")


def main():
    pasta_entrada = Path(PASTA_ENTRADA)
    pasta_saida = Path(PASTA_SAIDA)
    pasta_saida.mkdir(parents=True, exist_ok=True)

    arquivos = sorted(pasta_entrada.glob("*.xlsx"))

    if not arquivos:
        print("Nenhum arquivo .xlsx encontrado.")
        return

    print(f"Total de arquivos encontrados: {len(arquivos)}")

    for caminho_xlsx in arquivos:
        try:
            processar_arquivo(caminho_xlsx, pasta_saida)
        except Exception as e:
            print(f"ERRO em {caminho_xlsx.name}: {e}")

    print("Concluído.")


if __name__ == "__main__":
    main()