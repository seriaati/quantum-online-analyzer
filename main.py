import asyncio
import io
import os
import discord
from discord.ext import commands

from dotenv import load_dotenv
import pandas as pd

load_dotenv()


def analyze_excel(file: bytes, days: int) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file))

    filtered = df[["tablescraper-selected-row 3", "tablescraper-selected-row 10"]]
    filtered.columns = ["Code", "Date"]
    filtered = filtered[filtered["Code"].notna()]
    filtered = filtered[filtered["Date"].notna()]
    filtered = filtered[filtered["Code"] != "Security Description"]
    filtered = filtered[filtered["Date"] != "Call Date"]
    filtered["Date"] = filtered["Date"].str.replace(r"\s+", " ", regex=True).str.strip()
    filtered["Date"] = filtered["Date"].str.replace("n.a.", "").str.strip()
    filtered["Date"] = pd.to_datetime(filtered["Date"], errors="coerce")
    filtered = filtered[filtered["Date"].notna()]

    filtered = filtered[filtered["Date"] >= pd.Timestamp.now()]
    filtered = filtered[
        filtered["Date"] - pd.Timestamp.now() <= pd.Timedelta(days=days)
    ]
    return filtered


intents = discord.Intents.default()
bot = commands.Bot(command_prefix=commands.when_mentioned, intents=intents)


@bot.tree.command(name="分析", description="獲取距離今天指定天數內的資料")
@discord.app_commands.rename(file="檔案", days="天數")
@discord.app_commands.describe(file="要分析的 Excel 檔案", days="要取距離今天的天數")
async def analyze(
    i: discord.Interaction, file: discord.Attachment, days: int = 500
) -> None:
    await i.response.defer()
    df = await asyncio.to_thread(analyze_excel, await file.read(), days)
    if df.empty:
        await i.response.send_message("沒有找到符合條件的資料。", ephemeral=True)
        return

    df = df.sort_values("Date")
    df = df.reset_index(drop=True)
    df_str = df.to_string(index=False, header=False)
    await i.followup.send(content=df_str)


@commands.is_owner()
@bot.command()
async def sync(ctx: commands.Context) -> None:
    await bot.tree.sync()
    await ctx.send("Commands synced successfully!")


bot.run(os.getenv("TOKEN"))
