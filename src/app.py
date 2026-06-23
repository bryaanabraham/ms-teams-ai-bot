import asyncio
import re

from azure.identity import ManagedIdentityCredential
from microsoft_teams.api import MessageActivity, TypingActivityInput
from microsoft_teams.apps import ActivityContext, App
from config import Config

config = Config()

def create_token_factory():
    def get_token(scopes, tenant_id=None):
        credential = ManagedIdentityCredential(client_id=config.APP_ID)
        if isinstance(scopes, str):
            scopes_list = [scopes]
        else:
            scopes_list = scopes
        token = credential.get_token(*scopes_list)
        return token.token
    return get_token

app = App(
    token=create_token_factory() if config.APP_TYPE == "UserAssignedMsi" else None,
    skip_auth=not config.APP_ID,
)

@app.on_message
async def handle_message(ctx: ActivityContext[MessageActivity]):
    """Handle message activities using the new generated handler system."""
    await ctx.reply(TypingActivityInput())
    query = ctx.activity.text
    print(f"User entered: {query}")
    print("Calling LLM...")
    await ctx.send("Hey! The BrowsEZ teams bot is currently down. Please use the playground at https://app.browsez.in/playground for your queries.")
    # await ctx.send("Test payload recieved. Conversation ID is mapped to the backend.")

if __name__ == "__main__":
    asyncio.run(app.start())
