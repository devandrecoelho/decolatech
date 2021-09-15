// JavaScript source code
function CopyRequest(executionContext) {

    var quoteId = executionContext.data.entity.getId().replace("{", "").replace("}", "");
    var formContext = typeof executionContext.getFormContext != 'undefined' ? executionContext.getFormContext() : executionContext;
    var clientUrl = formContext.context.getClientUrl();
    formContext.ui.setFormNotification("Aguarde Copiando Cotação...", "WARNING", "flag");

    var data = {
        "quoteId": quoteId
    };

    var req = new XMLHttpRequest();

    req.open("POST", encodeURI(clientUrl + "/api/data/v9.2/quotes(" + quoteId + ")/Microsoft.Dynamics.CRM.dio_copyquote"), true);
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");

    req.onreadystatechange = function () {
        if (this.readyState === 4) {

            req.onreadystatechange = null;

            if (this.status === 200 || this.status === 204) {

                var result = JSON.parse(this.response);
                formContext.ui.clearFormNotification("flag");

                var sucessAlert = {
                    confirmButtonLabel: "Ok", text: "Cotação Copiada com sucesso !"
                };

                Xrm.Navigation.openAlertDialog(sucessAlert).then(function () {
                    var options = {
                        "entityName": "quote",
                        "entityId": result.copyQuoteId
                    };
                    Xrm.Navigation.openForm(options);
                });

            }
            else {

                var errorAlert = {
                    confirmButtonLabel: "Ok", text: "Erro ao Copiar Cotação" + "\n" + this.response
                };

                Xrm.Navigation.openAlertDialog(errorAlert).then(function () {
                });

            }
        }
    };

    req.send(JSON.stringify(data));

}