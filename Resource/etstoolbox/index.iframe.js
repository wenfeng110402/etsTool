console.log('test');
document.addEventListener('keydown', e => {
    console.log(e.key, e.key == 'F1');
    parent.postMessage('show-dialog', parent.location.origin);
});
