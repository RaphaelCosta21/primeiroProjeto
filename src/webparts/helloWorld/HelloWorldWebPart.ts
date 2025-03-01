import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

export interface IHelloWorldWebPartProps {
}
//teste para verificar se atualiza nos dois computadores
export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  public render(): void {
    // Definindo a estrutura básica no WebPart
    this.domElement.innerHTML = `
    <div class="webPartContainer">
      <h1>Bem-vindo ao WebPart!</h1>
      <div class="container">
        <button class="button" id="openFavoritesButton">Abrir Favoritos</button>
      </div>

      <!-- Pop-up Favoritos -->
      <div class="popup" id="favoritesPopup" style="display: none;">
        <div class="popupContent">
          <span class="closeButton" id="closePopup">&times;</span>
          <h2>Favoritos</h2>
          
          <!-- Tabela com abas -->
          <div class="tabContainer">
            <button class="tabButton" id="tab1Button">Aba 1</button>
            <button class="tabButton" id="tab2Button">Aba 2</button>
          </div>
          
          <div class="tabContent" id="tab1Content">
            <table>
              <thead>
                <tr>
                  <th>Item</th>
                  <th>Descrição</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td>Favorito 1</td>
                  <td>Descrição do Favorito 1</td>
                </tr>
                <tr>
                  <td>Favorito 2</td>
                  <td>Descrição do Favorito 2</td>
                </tr>
              </tbody>
            </table>
          </div>
          
          <div class="tabContent" id="tab2Content" style="display: none;">
            <table>
              <thead>
                <tr>
                  <th>Item</th>
                  <th>Descrição</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td>Favorito A</td>
                  <td>Descrição do Favorito A</td>
                </tr>
                <tr>
                  <td>Favorito B</td>
                  <td>Descrição do Favorito B</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  `;

    // Eventos de clique para abrir e fechar o pop-up
    this.domElement.querySelector("#openFavoritesButton")?.addEventListener("click", () => {
      this.showFavoritesPopup();
    });

    this.domElement.querySelector("#closePopup")?.addEventListener("click", () => {
      this.closeFavoritesPopup();
    });

    // Eventos para alternar entre as abas
    this.domElement.querySelector("#tab1Button")?.addEventListener("click", () => {
      this.showTabContent(1);
    });
    this.domElement.querySelector("#tab2Button")?.addEventListener("click", () => {
      this.showTabContent(2);
    });
  }

  private showFavoritesPopup(): void {
    const popup = this.domElement.querySelector("#favoritesPopup");
    if (popup) {
      (popup as HTMLElement).style.display = "flex";  // Exibe o pop-up
    }

    this.showTabContent(1); // Exibe a primeira aba ao abrir
  }

  private closeFavoritesPopup(): void {
    const popup = this.domElement.querySelector("#favoritesPopup");
    if (popup) {
      (popup as HTMLElement).style.display = "none";  // Esconde o pop-up
    }
  }

  private showTabContent(tabNumber: number): void {
    const allTabs = this.domElement.querySelectorAll(".tabContent");
    allTabs.forEach(tab => {
      (tab as HTMLElement).style.display = "none";
    });

    const tabContent = this.domElement.querySelector(`#tab${tabNumber}Content`);
    if (tabContent) {
      (tabContent as HTMLElement).style.display = "block";
    }
  }

  protected async onInit(): Promise<void> {
    return super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

//Função para criar o Excel e formatá-lo
  // private async createExcel(): Promise<void> {
  //   const workbook = new ExcelJS.Workbook();

  //   // Criando duas abas (worksheets) com nomes diferentes
  //   const primeiraAba = workbook.addWorksheet('CTR-Exported', {
  //     pageSetup: { paperSize: 9, orientation: 'portrait' }
  //   });

  //   const segundaAba = workbook.addWorksheet('Comentários Revisores', {
  //     pageSetup: { paperSize: 9, orientation: 'portrait' }
  //   });
        
  //   primeiraAba.unprotect();
  //   segundaAba.unprotect();

  //   workbook.creator = 'Raphael';
  //   workbook.calcProperties.fullCalcOnLoad = true;
  // //configurando a primeira aba
  //   primeiraAba.pageSetup.margins = {
  //     left: 0.14, right: 0.14,
  //     top: 0.14, bottom: 0.14,
  //     header: 0, footer: 0.15
  //   };
  //   primeiraAba.pageSetup.fitToPage = true;
  //   primeiraAba.pageSetup.fitToWidth = 1;
  //   primeiraAba.pageSetup.fitToHeight = 0;
  
  //   //configurando a segunda aba
  //   segundaAba.pageSetup.margins = {
  //     left: 0.14, right: 0.14,
  //     top: 0.14, bottom: 0.14,
  //     header: 0, footer: 0.15
  //   };
  //   segundaAba.pageSetup.fitToPage = true;
  //   segundaAba.pageSetup.fitToWidth = 1;
  //   segundaAba.pageSetup.fitToHeight = 0;  
  
  //   segundaAba.getColumn('A').width = 22;
  //   segundaAba.getColumn('B').width = 22;
  //   this.darkcell(segundaAba.getCell('A1'),'Comentários Revisores');
  //   segundaAba.mergeCells('A1:B1');
  //   this.normalcell(segundaAba.getCell('A2'),'Revisor 1');
  //   this.normalcell(segundaAba.getCell('A3'),'Revisor 2');
  //   await this.saveXLSX(workbook, "Criando Aba Revisor");
  // }
//Funções para editar as células
  // private darkcell(cell: ExcelJS.Cell,text: string, color = ''): void {
  //   cell.value = text;
  //   cell.alignment = { 
  //     vertical: 'middle', 
  //     horizontal: 'center',
  //     wrapText: true
  //   };
  //   cell.fill = {
  //     type: 'pattern',
  //     pattern:'solid',
  //     fgColor:{argb:(color !=='') ? color : '003B5C'},
  //   };
  //   cell.border = {
  //     top: {style:'thin'},
  //     left: {style:'thin'},
  //     bottom: {style:'thin'},
  //     right: {style:'thin'}
  //   };
  //   cell.font = {
  //     name: 'Calibri',
  //     bold: true,
  //     size: 12,
  //     color: {argb: "ffffff"}
  //   };
  // }
  // private normalcell(cell: ExcelJS.Cell,text: string, color = ''): void {
  //   cell.value = text;
  //   cell.alignment = { 
  //     vertical: 'middle', 
  //     horizontal: 'center',
  //     wrapText: true
  //   };
  //   cell.border = {
  //     top: {style:'thin'},
  //     left: {style:'thin'},
  //     bottom: {style:'thin'},
  //     right: {style:'thin'}
  //   };
  //   cell.font = {
  //     name: 'Calibri',
  //     bold: true,
  //     size: 12,
  //     color: {argb: "000000"}
  //   };
  // }
//Função para salvar o Excel
  // private saveXLSX = async (workbook: ExcelJS.Workbook, title: string): Promise<void> => {
  //   try {
  //     const xls64 = await workbook.xlsx.writeBuffer();
  //     // build anchor tag and attach file (works in chrome)
  //     const data = new Blob([xls64], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  //     const url = URL.createObjectURL(data);
  //     const a = document.createElement("a");
  //     a.href = url;
  //     a.download = `${title}.xlsx`;
  //     document.body.appendChild(a);
  //     a.click();
  //     setTimeout(() => {
  //       document.body.removeChild(a);
  //       window.URL.revokeObjectURL(url);
  //     }, 0);
  //     //loading.style.display = "none";
  //     //togglePopup ('popup','open',"Report Saved.")
  //     //saved = true;
  //   } catch (error) {
  //     console.log((error as Error).message);
  //     //loading.style.display = "none";
  //     //togglePopup ('popup','open',"ERROR. Not possible to save report.")
  //     //saved = true;
  //   }
  // }
}
