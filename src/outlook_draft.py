      import logging
      import asyncio
      import aiohttp
      from transformers import GPT4Model, GPT4Tokenizer

      tokenizer = GPT4Tokenizer.from_pretrained('gpt-4')
      model = GPT4Model.from_pretrained('gpt-4')

      async def get_unread_emails(session):
          async with session.get('https://outlook.office.com/api/v2.0/me/mailfolders/inbox/messages?$filter=IsRead eq false') as resp:
              return await resp.json()

      async def generate_draft_reply(session, email):
          try:
              # Use GPT-4 to generate a response
              inputs = tokenizer.encode(email['Body']['Content'], return_tensors='pt')
              outputs = model.generate(inputs, max_length=500, num_return_sequences=1)
              reply = tokenizer.decode(outputs[0], skip_special_tokens=True)
              return reply
          except Exception as e:
              logging.error(f"Failed to generate draft reply due to {str(e)}")

      async def save_draft_reply(session, draft_content, email):
          draft = {
              'Subject': f"RE: {email['Subject']}",
              'Body': {
                  'ContentType': 'Text',
                  'Content': draft_content
              },
              'ToRecipients': [
                  {
                      'EmailAddress': {
                          'Address': email['From']['EmailAddress']['Address']
                      }
                  }
              ]
          }
          await session.post('https://outlook.office.com/api/v2.0/me/messages', json=draft)

      async def process_email(session, email):
          try:
              draft_content = await generate_draft_reply(session, email)
              await save_draft_reply(session, draft_content, email)
          except Exception as e:
              logging.error(f"Failed to process email due to {str(e)}")