function teste(){
  try{
    resetarTodasVariaveis();
  }
  catch (erro) {  
    registrarErro(erro,true);
  }
}
function resetarTodasVariaveis(){
  const propriedades = PropertiesService.getDocumentProperties();
  propriedades.deleteAllProperties();  
}
