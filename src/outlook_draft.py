async def main():
    async with aiohttp.ClientSession() as session:
        emails = await get_unread_emails(session)
        if emails is not None:
            tasks = []
            for email in emails:
                task = process_email(session, email)
                tasks.append(task)
            await asyncio.gather(*tasks)

if __name__ == "__main__":
    asyncio.run(main())