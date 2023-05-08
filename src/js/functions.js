async function getBearerToken(login, password) {
  
  const apiUrl = 'https://gateway.fidelize.com.br/graphql'
  const query = `    
    mutation createToken {
      createToken(login: "${login}", password: "${password}") {
        token
      }
    }       
  `

  try {
    const response = await axios.post(apiUrl, { query }, {
      headers: {
        'Content-Type': 'application/json',        
      }
    })    

    return response.data.data.createToken.token
  } catch (error) {
    console.error('Erro na requisição:', error)
    document.getElementById('loading').style.display = 'none'; // Mostra o load enquanto a planilha carrega
    document.getElementById('box').style.display = 'flex'; // Esconde a div
    document.getElementById('root').style.display = 'flex'; // Exibe a div de erro    
  }
}

async function getOrderStatus(portal, order_code, bearer_token) {
  const apiUrl = 'https://gateway.fidelize.com.br/graphql'

  const query = `    
    {
      order(industry_code: "${portal}", order_code: ${order_code}) {
        industry_code
        order_code
        order_code
        industry_code
        wholesaler_code
        wholesaler_branch_code
        customer_code     
        status
        responses {
          id
          importation_outcome
          consideration_code
          imported_at
        }
        invoices {
          id
          importation_status
          importation_outcome
          status
          processed_at
          released_on
          danfe_key
          code
          value
          discount
          products_total_value
        }
      }
    }       
  `

  try {
    const response = await axios.post(apiUrl, { query }, {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${bearer_token}`
      }
    })    

    return response.data
  } catch (error) {
    console.error('Erro na requisição:', error)
  }
}

function formatPortal(portal) {    
  switch(portal) {
    case 'Abbott':
      portal = 'ABT'
      break
    case 'Astra Zeneca':
      portal = 'AZN'
      break
    case 'Exeltis':
      portal = 'EXE'
      break
    case 'Pfizer':
      portal = 'PFZ'
      break
    case 'webb':
      portal = 'SAN'
      break
    case 'RB':
      portal = 'RCK'
      break
    case 'Esanofi':
      portal = 'ESA'
      break
    case 'Baldacci':
      portal = 'BAL'
      break
    case 'Apsen':
      portal = 'APS'
      break
    case 'GSK': 
      portal = 'GSK'
      break
    case 'Servier':
      portal = 'SVR'
      break
    case 'Theraskin':
      portal = 'TKN'
      break
    case 'Upjhon':
      portal = 'UPJ'
      break
    case 'Bayer':
      portal = 'BAY'
      break
    default:
      portal = 'NÃO LOCALIZADO'
      break
  }
  
  return portal
}

function formatDate(date) {  
  let new_date = date.toLocaleString().split(',')      
  return new_date[0].replace(/\//g, '-')  
}

function formatStatus(status) {
  switch(status) {
    case 'INVOICE_RECEIVED':
    case 'INVOICE_PARTIALLY_CANCELED':
      status = 'ENVIADO - ESPELHO'
      break
    case 'CANCELLED':
    case 'CANCELLING':    
      status = 'CANCELADO'
      break
    case 'REJECTED':
      status = 'REJEITADO'
    default:
      status
      break    
  }

  return status
}

function createStatusColumn(worksheet) {
  const lastColumn = XLSX.utils.decode_range(worksheet['!ref']).e.c
  const newColumn = lastColumn + 1     
  const originColumn = XLSX.utils.encode_cell({ r: 0, c: newColumn })
  const columnTitle = 'STATUS SERVIMED'

  XLSX.utils.sheet_add_aoa(worksheet, [[columnTitle]], {origin: originColumn })
}

function checkSystemError(response) {
  let status

  if(response.errors) { // Trata se a api retornar erro. (quando o portal é "Produto sem ol")
    status = 'NÃO LOCALIZADO PARA O PORTAL'
  } else {
    if(response.data.status == 'AWAITING_RESPONSE' && response.data.order.responses[0] && response.data.order.responses[0].importation_outcome == 'Erro do sistema') {
      status = 'ERRO DO SISTEMA' // Se o pedido está aguardando retorno -> e possui erro de sistema
    } else if(response.data.status == 'AWAITING_INVOICE' && response.data.order.invoices[0] && response.data.order.invoices[0].importation_outcome == 'Erro do sistema') {
      status = 'ERRO DO SISTEMA' // Se o pedido está aguardando nota -> e possui erro de sistema
    } else {
      status = formatStatus(response.data.order.status) // Formata o status da api de ingles para portugues     
    }
  }

  return status
}

async function createStatusRows(worksheet, row, portal, order_code, bearer_token) {
  const lastColumnStatus = XLSX.utils.decode_range(worksheet['!ref']).e.c // Pega o número de referencia da ultima coluna da planilha
  const newColumnStatus = lastColumnStatus // Define em outra variavel para manter o padrão
  const originColumnStatus = XLSX.utils.encode_cell({ r: row+1, c: newColumnStatus }) // Transformar o numero de referencia da coluna para o valor da célula. (ex: A1, A2, ...)             
  
  const response = await getOrderStatus(portal, order_code, bearer_token) // Consulta a api        
  let orderStatus = checkSystemError(response)                            
  
  XLSX.utils.sheet_add_aoa(worksheet, [[orderStatus]], {origin: originColumnStatus }) // Adiciona os valores na coluna de status servimed 
}

let date = new Date()
date = formatDate(date)

const sendFile = document.getElementById('sendFile')
const form = document.getElementById('form')

form.addEventListener('submit', async (e) => { // Lê o arquivo selecionado quando houver mudança no file de arquivo
  e.preventDefault() 
  const login = document.getElementById('login').value
  const password = document.getElementById('password').value  
 
  document.getElementById('loading').style.display = 'flex'; // Mostra o load enquanto a planilha carrega
  document.getElementById('box').style.display = 'none'; // Esconde a div

  const file = sendFile.files[0]; // Localiza o arquivo
  const reader = new FileReader(); // Le o arquivo
  const bearer_token = await getBearerToken(login, password) // Pega o token da api

  reader.onload = async function(e) { // Inicia leitura do arquivo      
    let data = e.target.result; // Pega as informações do arquivo    
    let workbook = XLSX.read(data); // Lê a planilha            

    const promises = workbook.SheetNames.map(async sheetName => { // Inicia as tratativas
      const worksheet = workbook.Sheets[sheetName] // Pega os valores de cada aba
      
      createStatusColumn(worksheet) // Cria a coluna STATUS SERVIMED em todas as abas (se tiver mais de uma)
      
      const sheetData = XLSX.utils.sheet_to_json(worksheet) // Transforma os dados em json para realizar o proximo map  
      
      await Promise.all(sheetData.map(async (data, row) => {      
        let portal = data.PORTAL
        const order_code = data['ID Sub-Pedido']

        portal = formatPortal(portal) // Formata o portal transformando texto para sigla        
        await createStatusRows(worksheet, row, portal, order_code, bearer_token) // Cria as linhas com os respectivos status na coluna de STATUS SERVIMED               
      }))             
    })        

    await Promise.all(promises) // Aguarda todas as promisses finalizarem
    await XLSX.writeFile(workbook, `SERVIMED-${date}.xlsx`) // Salva a planilha com as novas informações adicionadas. Linha + Coluna
    
    document.getElementById('loading').style.display = 'none'; // Remove o load
    document.getElementById('box').style.display = 'flex'; // Volta a div
  }

  reader.readAsArrayBuffer(file);   
})