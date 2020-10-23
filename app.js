const { ipcRenderer } = require('electron')
const swal = require('sweetalert');

function protect() {

  let args = {
    wb: null,
    hashes: null
  };

  try {
    args.wb = document.getElementById("sheetFile").files.item(0).path;
    args.hashes = document.getElementById("PassWordFile").files.item(0).path;
    ipcRenderer.send('protect-sheet', args)
    $('.modal').modal('show');
  } catch {
    if (args.wb === null) {
      swal({
        title: "Aviso",
        text: "Selecione a planilha!",
        icon: "warning"
      });
    } else if (args.hashes === null) {
      swal({
        title: "Aviso",
        text: "Selecione o arquivo de hashes!",
        icon: "warning"
      });
    }
  }

}

function unProtect() {
  let args = {
    wb: null,
  };

  try {
    args.wb = document.getElementById("sheetFile").files.item(0).path;
    ipcRenderer.send('unProtect-sheet', args)
    $('.modal').modal('show');
  } catch {
    if (args.wb === null) {
      swal({
        // title:"Aviso",
        text: "Selecione a planilha!",
        icon: "warning"
      });
    }
  }
}

ipcRenderer.on("succ", (evt, args) => {
  $('.modal').modal('hide');
  swal({
    title: "Sucesso",
    icon: "success"
  })
})

ipcRenderer.on("err", (evt, args) => {
  $('.modal').modal('hide');
  swal({
    title: "Error",
    message: args,
    icon: "error"
  })

})