(function () {
    function clearRecipientDefaults() {
        var blocks = document.querySelectorAll('.body-param__example, .model-example, .opblock-body');
        blocks.forEach(function (block) {
            var inputs = block.querySelectorAll('input[type="text"]');
            inputs.forEach(function (input) {
                if (!input || input.value !== 'string') return;

                var container = input.closest('tr, .parameter, .prop-row, .model');
                var text = (container ? container.textContent : '') || '';
                var normalized = text.toLowerCase();

                if (normalized.includes('to') || normalized.includes('cc') || normalized.includes('bcc')) {
                    input.value = '';
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                }
            });
        });
    }

    var tries = 0;
    var timer = setInterval(function () {
        clearRecipientDefaults();
        tries++;
        if (tries > 30) clearInterval(timer);
    }, 500);

    document.addEventListener('click', function () {
        setTimeout(clearRecipientDefaults, 100);
    });
})();
