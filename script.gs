// BOTÃO DE LIMPAR
function limpar() {

  auxilio.getRange("B115").setValue(0); // imprime o número de petianos consultados até então...  
  clear_petianos();
  disponibilidade.getRange(l_equipe,c_equipe).clearContent();

}
  
// BOTÃO DE SELECIONAR TODOS
function selecionarTodos() {

  limpar();
  let a = 118;

  for(let i = l_pet; i < l_pet + 8; i++) {
    for(let j = c_pet; j < c_pet + 9; j = j + 4){
      if(auxilio.getRange(a,9).getValue() !== ' ') {
        disponibilidade.getRange(i,j).setValue(auxilio.getRange(a,9).getValue());
      }
      a++;    
    }
  } 

}

// BOTÃO DE CONSULTA GERAL
function consultaGeral() {

  let ver = 0;

  for(let i = l_pet; i < l_pet + 8; i++) {
    for(let j = c_pet; j < c_pet + 9; j = j + 4) {
      if(disponibilidade.getRange(i,j).getValue() != '' || equipe != '') {
        ver = 1;
      }
    }
  }
  
  if(ver == 1) {
    if(equipe !== '') {
      if(Browser.msgBox("Deseja consultar a disponibilidade da equipe do(a) "+equipe+"?",Browser.Buttons.YES_NO) == 'yes') {
        consultarEquipe();
      }
    } else {
      if(Browser.msgBox("Deseja consultar a disponibilidade do(s) petiano(s) selecionado(s)?",Browser.Buttons.YES_NO) == 'yes') {
        consultarPetianos();
      }
    }
  } else {
    SpreadsheetApp.getUi().alert("Por favor, selecione uma equipe ou o(s) petiano(s) que você deseja consultar.");
  }

}

// REALIZA A CONSULTA GERAL DOS PETIANOS SELECIONADOS
function consultarPetianos() {

  clear_spots();
  clear_load();
  clear_esp();
  auxilio.getRange("B115").setValue(0); // imprime o número de petianos consultados até então...  
  let n = []; // ?
  let matriz = []; // Armazenar as tabelas de horários de todos os petianos selecionados
  let rep = 1; // Verificar se o petiano fois selecionado mais de uma vez (repetição)
  let sparkline = 1; // Contar o número de petianos consultados (ajuda no funcionamento da barra de carregamento)

  for(let i = 0; i < 16; i++) {
    matriz[i] = [];
    n[i] = [];
    for(let j = 0; j < 7; j++) {
      matriz[i][j] = 0;
      n[i][j] = 0;
    }
  }
  
  for(let i = l_pet; i < l_pet + 8; i++) {
    for(let j = c_pet; j < c_pet + 9; j = j + 4) { 

      let petiano = disponibilidade.getRange(i,j).getValue(); // petiano que está sendo consultado      

      if(petiano != '') {

        let emBranco = false; // verifica se o horário está em branco ou preenchido adequadamente
        auxilio.getRange("B115").setValue(sparkline); // imprime o número de petianos consultados até então na aba auxílio
        sparkline++;  

        // Verifica se houve repetição entre os petianos selecionados
        if(j < 10) {
          for(let u = i; u < l_pet + 8; u++) {
            for(let v = j + 4; v < c_pet + 9; v = v + 4) {
              if(petiano == disponibilidade.getRange(u,v).getValue()) {
                rep = 0;
              }
            }
          }
        } 
        else if(j == 10) {
          for(let u = i + 1; u < l_pet + 8; u++) {
            for(let v = c_pet; v < c_pet + 9; v = v + 4) {
              if(petiano == disponibilidade.getRange(u,v).getValue()) {
                rep = 0;
              }
            }
          }         
        }
        if(rep == 0) {
          clear_load();
          clear_spots();
          return(Browser.msgBox(`ERRO: ${petiano} teve seu nome inserido repetidamente. Por favor, insira apenas uma vez.`));
        }

        let index = 0;

        // Busca o petiano na aba Mural
        for(let a = 1; a < linhah; a++) {

          // Manipula o "Nome Completo" do petiano para ficar apenas "Nome + Sobrenome"
          var inicio = ""+horarios.getRange(a,2).getValue()+"";
          var meio = inicio.split(" ");
          var termino = meio[0]+" "+meio[meio.length-1];

          // Ao encontrar o petiano na aba Mural, armazena todos seus horários no array "matriz"
          if(petiano === termino) {

            index = 1;            
            let x = c_spots;
            let y = l_spots;
            let r = 0;
            let k = 0;

            for(let l = a; l <= a + 15; l++) {
              for(let c = c_slots; c <= c_slots + 6; c++) {

                let slot = horarios.getRange(l,c).getValue();

                switch(slot) {
                  case 'Disponível':
                      slot = 1;
                      break;
                  case 'Flexível':
                      slot = 1;
                      break;  
                  case 'Ocupado':
                      slot = 0;
                      break;
                  default:
                      slot = 0;
                      emBranco = true;
                      n[r][k]--;
                      break;
                }
                
                /*
                if(slot == 'Disponível') {
                  slot = 1;
                } else if(slot == 'Flexível') {
                  slot = 1/2;
                } else if(slot == 'Ocupado') {
                  slot = 0;
                } else {
                  slot = 0;                  
                  emBranco = true;
                  n[r][k] = n[r][k] - 1                  
                }
                */

                matriz[r][k] = matriz[r][k]+slot;
                n[r][k] = n[r][k] + 1; 

                k = k + 1;
                x = x + 1;

              }
              
              x = 5;
              y = y + 1;
              k = 0;
              r = r + 1;

            }

            break;

          }
        }

        // Sinalizar alguns erros verificados
        if(emBranco === true) {
          Browser.msgBox("ERRO: "+petiano+" não preencheu o Mural de Horários adequadamente. O resultado da consulta será comprometido.")
        }
        else if(index !== 1) {
          Browser.msgBox("ERRO: "+petiano+" ainda não preencheu o Mural de Horários. O resultado da consulta não levará em consideração esse petiano.")
        }

      }
    }
  }

  // Calcula as disponibilidades dos petianos e imprimí-las nos slots de resultado da consulta geral
  
  let colSpot = c_spots;
  let linSpot = l_spots;  

  for(let i = 0; i < 16; i++) {
    for(let j = 0; j < 7; j++) {

      let result = ""+matriz[i][j]*100/n[i][j]+"";
      let sep = result.split(".");

      if(sep[1] != null) {
        disponibilidade.getRange(linSpot,colSpot).setFormula(`=IFERROR(ROUNDUP(${sep[0]},${sep[1]}%;2);"")`);
      } else {
        disponibilidade.getRange(linSpot,colSpot).setFormula(`=IFERROR(${result}%;"")`);
      }

      colSpot++;

    }

    linSpot++;
    colSpot -= 7;

  }
}

// REALIZA A CONSULTA GERAL DA EQUIPE SELECIONADA
function consultarEquipe() {

  let lTemp = l_pet;
  let cTemp = c_pet;
  clear_petianos();
  clear_load();
  clear_spots();
  auxilio.getRange("B115").setValue(0); // imprime o número de petianos consultados até então...

  for(let i = 2; i < linhaa; i++) {

    if(equipe == auxilio.getRange(i,3).getValue()) {

      for(let j = 4; j < colunaa; j++) {
        
        if(auxilio.getRange(i,j).getValue() !== '') {

          disponibilidade.getRange(lTemp,cTemp).setValue(auxilio.getRange(i,j).getValue());

          if(cTemp <= 6) {
            cTemp = cTemp + 4;
          } else if(cTemp > 6) {
            cTemp = 2;
            lTemp = lTemp + 1;
          }

        }
      }
    }
  }

  consultarPetianos();

}

// BOTÃO DE CONSULTA ESPECÍFICA
function consultaEspecifica() {

  // Limpa os campos

  auxilio.getRange("B115").setValue(0); // limpa a barra de carregamento
  disponibilidade.getRange('B52:M59').clearContent();
  disponibilidade.getRange('B52:M59').setBackground(null);  

  // Verifica os erros
  
  let consEsp = disponibilidade.getRange("B45:M45").getValue(); // horário consultado

  if(disponibilidade.getRange(l_spots,c_spots).getValue() == '' && disponibilidade.getRange(l_pet,c_pet).getValue() == '') {
    return Browser.msgBox("Para realizar a consulta específica, antes é necessário realizar a consulta geral.");
  } 
  else if(consEsp == '') {
    return Browser.msgBox("Por favor, selecione o horário que você deseja consultar.");
  }  

  // Consulta...

  let sparkline = 1; // barra de carregamento, conta os petianos consultados até então...
  let consEspL,
      consEspC = parseInt(consEsp[0]);
  let resultEsp = []; // vetor que vai armazenar os resultados da consulta
  
  for(let place = 0; place < 24; place++) { // transforma "resultEsp" em uma matriz de 24 linhas
    resultEsp[place] = [];
  }

  // Manipula o horário consultado para ser possível realizar a consulta na aba Mural
  switch(consEsp[1]) {
    case "M":
        for(let j = 0; j < 6; j++) {
          if(consEsp[2] == j+1) {
            consEspL = j;
          }
        }
        break;
    case "T":
        for(let j = 0; j < 6; j++) {
          if(consEsp[2] == j+1) {
            consEspL = j+6;
          }
        }
        break;
    case "N":
        for(let j = 0; j < 4; j++) {
          if(consEsp[2] == j+1) {
            consEspL = j+12;
          }
        }
        break;               
  }
  
  let petiano, 
      nomeCompleto, 
      nomeSobrenome,
      slotEsp, // armazena o slot do horário consultado
      pos = 0; // percorre as posições do vetor "resultEsp"

  // Busca o horário de cada petiano na aba Mural e imprime na aba Consulta os resultados
  for(let i = l_pet; i < l_pet + 8; i++) {
    for(let j = c_pet; j < c_pet + 9; j = j + 4) {

      petiano = String(disponibilidade.getRange(i,j).getValue());

      // Informações mínimas usadas para alocar o petiano no resultado da consulta específica
      resultEsp[pos][0] = i+43;
      resultEsp[pos][1] = j;
    //resultEsp[pos][2] = null;
      resultEsp[pos][3] = petiano;

      if(petiano != '') {

        for(let a = 1; a < linhah; a++) {

          nomeCompleto = String(horarios.getRange(a,2).getValue()).split(" ");
          nomeSobrenome = `${nomeCompleto[0]} ${nomeCompleto[nomeCompleto.length-1]}`;    

          if(nomeSobrenome ===  petiano) {  

            auxilio.getRange("B115").setValue(sparkline); // imprime o número de petianos consultados até então...
            sparkline++;   

            slotEsp = horarios.getRange((a+consEspL),(11+consEspC)).getValue();  

            switch(slotEsp) {
              case 'Disponível':
                  resultEsp[pos][2] = '#64e678';
                  break;
              case 'Flexível':
                  resultEsp[pos][2] = '#f5cc52';
                  break;  
              case 'Ocupado':
                  resultEsp[pos][2] = '#eb6262';
                  break;
              default:
                  resultEsp[pos][2] = null; 
                  break;
            }

            break; // QUE SACADA INCRÍVEL! ECONOMIZA CERCA DE 30 SEGUNDOS DO TEMPO DE CARREGAMENTO! PQ N PENSEI ANTES?

          }                
        }
      }

      pos++;

    }
  }

  // Imprime o resultado na planilha
  for(let pos in resultEsp) {
      disponibilidade
        .getRange(resultEsp[pos][0],resultEsp[pos][1])
        .setBackground(resultEsp[pos][2])
        .setValue(resultEsp[pos][3]);
  }

}

// BOTÃO DE FINALIZAR
function finalizarConsulta() {

  clear_consultas();
  Browser.msgBox("Consulta finalizada.");

}