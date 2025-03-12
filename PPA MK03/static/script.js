document.addEventListener('DOMContentLoaded', function() {
    function openTab(evt, tabName) {
        var i, tabcontent, tablinks;
        tabcontent = document.getElementsByClassName("tabcontent");
        for (i = 0; i < tabcontent.length; i++) {
            tabcontent[i].style.display = "none";
        }

        tablinks = document.getElementsByClassName("tablinks");
        for (i = 0; i < tablinks.length; i++) {
            tablinks[i].className = tablinks[i].className.replace(" active", "");
        }

        document.getElementById(tabName).style.display = "block";
        if (evt != null) {
            evt.currentTarget.className += " active";
        }
    }

    // Adiciona event listeners para os campos "Outros" em todas as abas
    const outros = document.getElementById("outros");
    if (outros) {
        outros.addEventListener("change", function() {
            let outrosInput = document.getElementById("outrosInput");
            outrosInput.style.display = this.checked ? "block" : "none";
        });
    }

    const outrosEpi = document.getElementById("outros_epi");
    if (outrosEpi) {
        outrosEpi.addEventListener("change", function() {
            let outrosInput = document.getElementById("outros_epi_input");
            outrosInput.style.display = this.checked ? "block" : "none";
        });
    }

    const outrosEquipamentos = document.getElementById("outros_equipamentos");
    if (outrosEquipamentos) {
        outrosEquipamentos.addEventListener("change", function() {
            let outrosInput = document.getElementById("outrosInput_equipamentos");
            outrosInput.style.display = this.checked ? "block" : "none";
        });
    }

    // Nova lógica para navegação entre abas com validação
    const nextTabButtons = document.querySelectorAll('.next-tab');
    nextTabButtons.forEach(button => {
        button.addEventListener('click', () => {
            const currentTab = button.parentElement;
            const currentTabId = currentTab.id;
            const nextTabId = getNextTabId(currentTabId);
            if (nextTabId) {
                // Removida a validação
                openTab(null, nextTabId);
            }
        });
    });

    function getNextTabId(currentTabId) {
        const tabIds = ['dados_essenciais', 'stop6', 'zero6', 'perigos_riscos', 'epis', 'equipamentos', 'finalizacao'];
        const currentTabIndex = tabIds.indexOf(currentTabId);
        if (currentTabIndex < tabIds.length - 1) {
            return tabIds[currentTabIndex + 1];
        }
        return null;
    }

    const tablinks = document.querySelectorAll('.tablinks');
    tablinks.forEach(tablink => {
        tablink.addEventListener('click', (event) => {
            const tabId = tablink.dataset.tab;
            // Removida a validação
            openTab(event, tabId);
        });
    });

    // Abre a aba "dados_essenciais" por padrão
    openTab(null, 'dados_essenciais');

    const gerarAtep = document.getElementById('gerar-atep');
    if (gerarAtep) {
        gerarAtep.addEventListener('click', () => {
            document.querySelector('form').submit(); // Envia o formulário para o backend
        });
    }
});