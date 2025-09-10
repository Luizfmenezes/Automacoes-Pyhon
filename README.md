1. orquestrador.py
É o programa principal com a interface gráfica. É a "central de comando" onde você digita as datas, escolhe quais robôs quer rodar (ICF, ICV, IPP) e clica no botão para iniciar tudo. Ele mostra o que está acontecendo em tempo real.

2. auto-python/automacao_icf.py
É um robô focado em uma única tarefa: baixar o relatório ICF do sistema da SPTrans. Ele faz o login, navega até a tela certa, baixa o arquivo e manda o Excel importar os dados.

3. auto-python/automacao_icv_icvfh.py
É o robô que baixa dois relatórios de uma vez: o ICV (mais detalhado) e o ICVFH (resumido por hora). Ele entra no sistema, preenche a data e baixa os dois arquivos, um depois do outro, e manda o Excel organizá-los.

4. auto-python/automacao_ipp.py
Este é o robô que cuida dos relatórios de pontualidade, o IPP e o IPPFH. Ele é um pouco mais complexo porque baixa esses relatórios para cada grupo de linhas (D1 e D2), um de cada vez. No final, ele avisa o Excel para juntar todas as informações.

5. analise_diaria.py
Esta é a sua ferramenta de análise. Depois que os robôs baixaram e organizaram os dados, você usa este programa para gerar aquele resumo diário formatado. Você informa o dia, escolhe as linhas e ele cria o texto pronto para ser copiado.

6. auto-nimer/nimer_scrap.py e nimer_scrap_D2.py
São robôs "fotógrafos". Eles entram no sistema Nimer, pesquisam a data que você pediu e "tiram uma foto" dos dados de pendências e fotos das linhas. Em vez de uma planilha, eles criam uma imagem .png que mostra o resultado de forma bem visual, como um gráfico. Os dois arquivos fazem a mesma coisa, mas para grupos de linhas diferentes (D1 e D2).
