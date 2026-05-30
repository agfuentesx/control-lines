function onOpen() {
  FormApp.getUi().createMenu("PS Functions").addItem("Reset the form", "resetForm").addToUi()
}

function resetForm() {
  let form=FormApp.getActiveForm()
  let items=form.getItems()
  items.forEach(item=>form.deleteItem(item))
  form.deleteAllResponses()
}
