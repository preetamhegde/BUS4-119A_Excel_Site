function SubmitButton(){
    console.log('working');
    var textareaValue = document.getElementById("datainput").value;
    fetch("https://excelsite-cf81471dcbae.herokuapp.com/update_excel", {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ data: textareaValue })
        
    })

    .then(response => {
        if (!response.ok) {
            throw new Error('Network response was not ok ' + response.statusText);
        }
        return response.json();
    })
    .then(data => {
        document.querySelector('.excel-container').innerHTML = data.updatedTable;
    })
    .catch(error => {
        console.error('There has been a problem with your fetch operation:', error);
    });
}



function ResetButton(){
    console.log('0');
    document.getElementById("datainput").value = "";
}



function bindEventListenersForPage(){
    var submitButton = document.getElementById("Submit-Button");

    if(submitButton){
        submitButton.removeEventListener("click", SubmitButton);
        submitButton.addEventListener("click", SubmitButton);
    }
    var resetButton = document.getElementById("Reset-Button");

    if(resetButton){
        resetButton.removeEventListener("click", ResetButton);
        resetButton.addEventListener("click", ResetButton);
    }
}



document.addEventListener("DOMContentLoaded", function() {
    bindEventListenersForPage();
});