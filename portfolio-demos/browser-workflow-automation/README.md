# Browser Workflow Automation Demo

A public JavaScript/HTML demo showing how repetitive browser workflows can be automated safely.

## What it demonstrates

The automation script:

- changes form values
- dispatches browser events
- submits a repeated workflow across several environments
- waits asynchronously for completion
- detects temporary failures
- retries failed attempts
- validates the final state
- produces an execution log

## Scenario

The included mock portal simulates a repetitive report-generation task across three fictional environments:

- North
- Central
- South

The **Central** environment intentionally returns a temporary error on the first attempt. The automation detects it and retries.

## Try it

1. Open `index.html` in a browser.
2. Open Developer Tools.
3. Open the Console.
4. Paste the contents of `automation.js`.
5. Review the live status panel and final `console.table` output.

## Why this matters

The point is not the fictional report portal. The transferable pattern is:

**configure → trigger → wait → validate → retry → log**

That same pattern appears in many browser-based administrative workflows.

## Data safety

This is an original mock application. It contains no private URLs, production selectors, customer data or employer-specific logic.
