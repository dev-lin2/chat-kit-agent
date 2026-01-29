# chat-kit-agent-sharepoint-webpart

## Summary

Webpart component for Open AI's Chat Kit Agent. To use the Chat Kit we need worflow id and a place where we can generate session token.
You can take a look at this sample lambda function below

## How to use
- replace `{tenantDomain}` in `config/serve.json` and `.vscode/launch.json` if needed
- run `npm run build`
- Follow this official documents how to add custom webparts to sharepoint below

## References

- [How to deploy](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/serve-your-web-part-in-a-sharepoint-page)
- [Sample Lambda function](https://github.com/dev-lin2/chat-kit-agent-token-generator)

