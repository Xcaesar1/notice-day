var cp = DebugSymbol.getFunctionByName('CreateProcessW');

Interceptor.attach(cp, {
  onEnter(args) {
    try {
      var cmd = args[1].readUtf16String();
      if (
        cmd &&
        cmd.indexOf('ziniaobrowser') !== -1 &&
        cmd.indexOf('--type=') === -1 &&
        cmd.indexOf('--remote-debugging-port') === -1
      ) {
        var newCmd = cmd + ' --remote-debugging-port=0';
        args[1] = Memory.allocUtf16String(newCmd);
        console.log('[ZINIAO_CDP] injected remote debugging port');
      }
    } catch (e) {
      console.log('[ZINIAO_CDP] hook error: ' + e);
    }
  },
});

console.log('[ZINIAO_CDP] hook active');
