import pandas as pd
import pathlib
import win32com as win32

#Read Excel Files.
emails = pd.read_excel('Emails.xlsx')
sales = pd.read_excel('Vendas.xlsx')
stores = pd.read_csv('Lojas.csv', encoding = 'latin1', sep = ';')
sales = sales.merge(stores, on = 'ID Loja')
stores_dict = {}
#Creating new Dataframe.
for store in stores['Loja']:
    stores_dict[store] = sales.loc[sales['Loja'] == store, :]
#Creating Backup files for each store.
day = sales['Data'].max()
backup_path = pathlib.Path(r'backup')
backup = backup_path.iterdir()
backupfile_list = [file.name for file in backup]
for store in stores_dict:
    if store not in backupfile_list:
        new_folder = backup_path / store
        new_folder.mkdir()
    file_name = f'{day.day}_{day.month}_{store}.xlsx'
    file_loc = backup_path / store / file_name
    stores_dict[store].to_excel(file_loc)
#Daily and Yearly Goals.
daily_yield_goal = 1000
yearly_yield_goal = 1650000
daily_prod_goal = 4
yearly_prod_goal = 120
avrg_ticket_goal = 500
#Sending the E-Mails for each Store Manager.
for store in stores_dict:
    store_sales = stores_dict[store]
    day_store_sales = store_sales.loc[store_sales['Data'] == day, :]
    year_yield = store_sales['Valor Final'].sum()
    day_yield = day_store_sales['Valor Final'].sum()
    prod_qnt = len(store_sales['Produto'].unique())
    day_prod_qnt = len(day_store_sales['Produto'].unique())
    sale_value = store_sales.groupby('Código Venda').sum()
    avrg_ticket = sale_value['Valor Final'].mean()
    day_sale_value = day_store_sales.groupby('Código Venda').sum()
    day_avrg_ticket = day_sale_value['Valor Final'].mean()
    outlook = win32.Dispatch('outlook.application')
    nome = emails.loc[emails['Loja'] == store, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja'] == store, 'E-mail'].values[0]
    mail.Subject = f'OnePage Dia {day.day}/{day.month} - Loja {store}'
    #Checks each if goals were reach, returns green if yes, else red.
    if day_yield >= daily_yield_goal:
        color_y_day = 'green'
    else:
        color_y_day = 'red'
    if year_yield >= yearly_yield_goal:
        color_y_year = 'green'
    else:
        color_y_year = 'red'
    if day_prod_qnt >= daily_prod_goal:
        color_q_day = 'green'
    else:
        color_q_day = 'red'
    if prod_qnt >= yearly_prod_goal:
        color_q_year = 'green'
    else:
        color_q_year = 'red'
    if day_avrg_ticket >= avrg_ticket_goal:
        color_t_day = 'green'
    else:
        color_t_day = 'red'
    if avrg_ticket >= avrg_ticket_goal:
        color_t_year = 'green'
    else:
        color_t_year = 'red'
    #Email HTML
    mail.HTMLBody = f'''
        <p>Bom dia, {nome}</p>

        <p>O resultado de ontem <strong>({day.day}/{day.month})</strong> da <strong>Loja {store}</strong> foi:</p>

        <table>
          <tr>
            <th>Indicador</th>
            <th>Valor Dia</th>
            <th>Meta Dia</th>
            <th>Cenário Dia</th>
          </tr>
          <tr>
            <td>Faturamento</td>
            <td style="text-align: center">R${day_yield:.2f}</td>
            <td style="text-align: center">R${daily_yield_goal:.2f}</td>
            <td style="text-align: center"><font color="{color_y_day}">◙</font></td>
          </tr>
          <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align: center">{day_prod_qnt}</td>
            <td style="text-align: center">{daily_prod_goal}</td>
            <td style="text-align: center"><font color="{color_q_day}">◙</font></td>
          </tr>
          <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${avrg_ticket:.2f}</td>
            <td style="text-align: center">R${avrg_ticket_goal:.2f}</td>
            <td style="text-align: center"><font color="{color_t_day}">◙</font></td>
          </tr>
        </table>
        <br>
        <table>
          <tr>
            <th>Indicador</th>
            <th>Valor Ano</th>
            <th>Meta Ano</th>
            <th>Cenário Ano</th>
          </tr>
          <tr>
            <td>Faturamento</td>
            <td style="text-align: center">R${year_yield:.2f}</td>
            <td style="text-align: center">R${yearly_yield_goal:.2f}</td>
            <td style="text-align: center"><font color="{color_y_year}">◙</font></td>
          </tr>
          <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align: center">{prod_qnt}</td>
            <td style="text-align: center">{yearly_prod_goal}</td>
            <td style="text-align: center"><font color="{color_q_year}">◙</font></td>
          </tr>
          <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${avrg_ticket:.2f}</td>
            <td style="text-align: center">R${avrg_ticket_goal:.2f}</td>
            <td style="text-align: center"><font color="{color_t_year}">◙</font></td>
          </tr>
        </table>

        <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

        <p>Qualquer dúvida estou à disposição.</p>
        <p>Att., Lira</p>
        '''
    attachment = pathlib.Path.cwd() / backup_path / store / f'{day.day}_{day.month}_{store}.xlsx'
    mail.Attachments.Add(str(attachment))
    mail.Send()
#Getting the highest and lowest yielding Stores for the Director
store_yield = sales.groupby('Loja')[['Loja', 'Valor Final']].sum()
year_store_yield = store_yield.sort_values(by = 'Valor Final', ascending = False)
ranking_file_name = f'{day.day}_{day.month}_Ranking_Anual.xlsx'
year_store_yield.to_excel(r'backup\{}'.format(ranking_file_name))
day_sales = sales[sales['Data'] == day, :]
day_store_yield = day_sales.groupby('Loja')[['Loja', 'Valor Final']].sum()
day_store_yield = day_store_yield.sort_values(by = 'Valor Final', ascending = False)
day_ranking_file_name = f'{day.day}_{day.month}_Ranking_Dia.xlsx'
day_store_yield.to_excel(r'backup\{}'.format(ranking_file_name))
#Sending E-mail to the Director
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']=='Diretoria', 'E-mail'].values[0]
mail.Subject = f'Ranking Dia {day.day}/{day.month}'
mail.Body = f'''
Prezados, bom dia

Melhor loja do Dia em Faturamento: Loja {day_store_yield.index[0]} com Faturamento R${day_store_yield.iloc[0, 0]:.2f}
Pior loja do Dia em Faturamento: Loja {day_store_yield.index[-1]} com Faturamento R${day_store_yield.iloc[-1, 0]:.2f}

Melhor loja do Ano em Faturamento: Loja {year_store_yield.index[0]} com Faturamento R${year_store_yield.iloc[0, 0]:.2f}
Pior loja do Ano em Faturamento: Loja {year_store_yield.index[-1]} com Faturamento R${year_store_yield.iloc[-1, 0]:.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição.

Att.,
Lira
'''
attachment = pathlib.Path.cwd() / backup_path / f'{day.month}_{day.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd() / backup_path / f'{day.month}_{day.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))
mail.Send()


