//---- Error Codes:
//----  [01] - Ajax call for dropdownlist data
//----  [02] - Ajax call for column model
//----  [03] - Ajax call for row data
//----  [04] - Ajax call for modal buttons
//----  [05] - Ajax call to generate ip address text file

//---- Set up global JS variables
var DDLPatchUser;

$(document).ready(function () {    
    if (document.getElementById("FileUpload1").value == "") {
        $("#tbFileName").val("");
    }
});

$("#FileUpload1").change(function () {
    if (document.getElementById("FileUpload1").value != "") {
        $("#cbNoDocReq").attr("checked", false);
        $("#cbNoDocReq").prop("checked", false);
        $("#cbNoDocReq").attr("disabled", true);
    }
    else {
        $("#tbFileName").val("");
    }
})

$("#ddProposedMajor").change(function () {
    $("#cbxUpdateCategory").attr("disabled", true);
    if ($("#ddProposedMajor option:selected").index() > 0) {        
        $("#ddProposedMinor").attr("disabled", false);
    }
    else {
        $("#ddProposedMinor").attr("disabled", true);
        $("#cbxUpdateCategory").attr("disabled", true);
    }
});

$("#ddProposedMinor").change(function () {
    if ($("#ddProposedMinor option:selected").index() > 0) {
        $("#cbxUpdateCategory").attr("disabled", false);
    }
    else {
        $("#cbxUpdateCategory").attr("disabled", true);
    }
});

$("#tbProposedSecured").blur(function calculateTotal() {
    var proposedSecured = $("#tbProposedSecured").val().replace("$", "").replace(",", "");
    var proposedAdmin = $("#tbProposedAdmin").val().replace("$", "").replace(",", "");
    var proposedPriority = $("#tbProposedPriority").val().replace("$", "").replace(",", "");
    var proposedUnsecured = $("#tbProposedUnsecured").val().replace("$", "").replace(",", "");
    var totalAmt = $("#tbProposedTotal").val().replace("$", "").replace(",", "");
    if (!isNaN(proposedSecured) && !isNaN(proposedAdmin) && !isNaN(proposedPriority) && !isNaN(proposedUnsecured)) {
        totalAmt = parseFloat(proposedAdmin) + parseFloat(proposedPriority) + parseFloat(proposedSecured) + parseFloat(proposedUnsecured);
    }
    
    $("#tbProposedTotal").val(totalAmt);
})
$("#tbProposedAdmin").blur(function calculateTotal() {
    var proposedSecured = $("#tbProposedSecured").val().replace("$", "").replace(",", "");
    var proposedAdmin = $("#tbProposedAdmin").val().replace("$", "").replace(",", "");
    var proposedPriority = $("#tbProposedPriority").val().replace("$", "").replace(",", "");
    var proposedUnsecured = $("#tbProposedUnsecured").val().replace("$", "").replace(",", "");
    var totalAmt = $("#tbProposedTotal").val().replace("$", "").replace(",", "");
    if (!isNaN(proposedSecured) && !isNaN(proposedAdmin) && !isNaN(proposedPriority) && !isNaN(proposedUnsecured)) {
        totalAmt = parseFloat(proposedAdmin) + parseFloat(proposedPriority) + parseFloat(proposedSecured) + parseFloat(proposedUnsecured);
    }

    $("#tbProposedTotal").val(totalAmt);
})
$("#tbProposedPriority").blur(function calculateTotal() {
    var proposedSecured = $("#tbProposedSecured").val().replace("$", "").replace(",", "");
    var proposedAdmin = $("#tbProposedAdmin").val().replace("$", "").replace(",", "");
    var proposedPriority = $("#tbProposedPriority").val().replace("$", "").replace(",", "");
    var proposedUnsecured = $("#tbProposedUnsecured").val().replace("$", "").replace(",", "");
    var totalAmt = $("#tbProposedTotal").val().replace("$", "").replace(",", "");
    if (!isNaN(proposedSecured) && !isNaN(proposedAdmin) && !isNaN(proposedPriority) && !isNaN(proposedUnsecured)) {
        totalAmt = parseFloat(proposedAdmin) + parseFloat(proposedPriority) + parseFloat(proposedSecured) + parseFloat(proposedUnsecured);
    }

    $("#tbProposedTotal").val(totalAmt);
})
$("#tbProposedUnsecured").blur(function calculateTotal() {
    var proposedSecured = $("#tbProposedSecured").val().replace("$", "").replace(",", "");
    var proposedAdmin = $("#tbProposedAdmin").val().replace("$", "").replace(",", "");
    var proposedPriority = $("#tbProposedPriority").val().replace("$", "").replace(",", "");
    var proposedUnsecured = $("#tbProposedUnsecured").val().replace("$", "").replace(",", "");
    var totalAmt = $("#tbProposedTotal").val().replace("$", "").replace(",", "");
    if (!isNaN(proposedSecured) && !isNaN(proposedAdmin) && !isNaN(proposedPriority) && !isNaN(proposedUnsecured)) {
        totalAmt = parseFloat(proposedAdmin) + parseFloat(proposedPriority) + parseFloat(proposedSecured) + parseFloat(proposedUnsecured);
    }

    $("#tbProposedTotal").val(totalAmt);
})
