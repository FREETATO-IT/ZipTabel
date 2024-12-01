window.inputChanged = function () {
    previousInput = null;
    
    const inputElements = document.querySelectorAll(".input-cell");
    
    inputElements.forEach(el => {
        el.readOnly = false; // Отключаем все, кроме текущего
    });
};