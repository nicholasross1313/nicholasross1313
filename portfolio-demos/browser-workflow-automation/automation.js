(async () => {
  const environments = ["North", "Central", "South"];
  const dateFrom = "2026-09-01";
  const dateTo = "2026-09-30";
  const maxAttempts = 3;

  const results = [];

  for (const environment of environments) {
    const result = await runEnvironment(environment, dateFrom, dateTo, maxAttempts);
    results.push(result);
  }

  console.table(results);

  async function runEnvironment(environment, from, to, limit) {
    for (let attempt = 1; attempt <= limit; attempt++) {
      setSelectValue("#environment", environment);
      setInputValue("#dateFrom", from);
      setInputValue("#dateTo", to);

      document.querySelector("#generateBtn").click();

      const state = await waitForTerminalState(5000);

      if (state === "success") {
        return {
          environment,
          status:"Successful",
          attempts:attempt,
          message:document.querySelector("#status").textContent
        };
      }

      if (state === "retryable" && attempt < limit) {
        console.warn(`${environment}: retrying after temporary error (attempt ${attempt})`);
        await delay(500);
        continue;
      }

      return {
        environment,
        status:"Failed",
        attempts:attempt,
        message:document.querySelector("#status").textContent
      };
    }
  }

  function setSelectValue(selector, value) {
    const element = document.querySelector(selector);
    if (!element) throw new Error(`Missing element: ${selector}`);

    element.value = value;
    element.dispatchEvent(new Event("change", {bubbles:true}));
  }

  function setInputValue(selector, value) {
    const element = document.querySelector(selector);
    if (!element) throw new Error(`Missing element: ${selector}`);

    element.value = value;
    element.dispatchEvent(new Event("input", {bubbles:true}));
    element.dispatchEvent(new Event("change", {bubbles:true}));
  }

  async function waitForTerminalState(timeoutMs) {
    const started = Date.now();

    while (Date.now() - started < timeoutMs) {
      const status = document.querySelector("#status");
      if (!status) throw new Error("Status element disappeared.");

      const state = status.dataset.state;

      if (["success","retryable","error"].includes(state)) {
        return state;
      }

      await delay(100);
    }

    return "timeout";
  }

  function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
})();
