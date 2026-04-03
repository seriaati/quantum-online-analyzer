import asyncio
import io
import os
from typing import Literal
import discord
from discord.ext import commands

from dotenv import load_dotenv
import pandas as pd
from typing import Optional

load_dotenv()


def analyze_excel(
    file: bytes,
    *,
    start_day: int = None,
    end_day: int = None,
) -> pd.DataFrame:
    if start_day > end_day:
        raise ValueError("起始天數不能大於結束天數")

    df = pd.read_excel(io.BytesIO(file))

    filtered = df[
        [
            "tablescraper-selected-row 3",
            "tablescraper-selected-row 10",
            "tablescraper-selected-row 4",
        ]
    ]

    filtered.columns = ["Code", "Date", "Security Description"]
    filtered = filtered[filtered["Code"].notna()]
    filtered = filtered[filtered["Date"].notna()]

    filtered = filtered[filtered["Code"] != "Security Description"]
    filtered = filtered[filtered["Date"] != "Call Date"]

    filtered["Date"] = filtered["Date"].str.replace("Call Date:", "").str.strip()
    filtered["Date"] = filtered["Date"].str.replace(r"\s+", " ", regex=True).str.strip()
    filtered["Date"] = filtered["Date"].str.replace("n.a.", "").str.strip()
    filtered["Date"] = filtered["Date"].str.replace("None", "").str.strip()

    filtered["Date"] = pd.to_datetime(filtered["Date"], errors="coerce")
    filtered = filtered[filtered["Date"].notna()]

    now = pd.Timestamp.now()
    start_date = now + pd.Timedelta(days=start_day)
    end_date = now + pd.Timedelta(days=end_day)

    filtered = filtered[
        (filtered["Date"] >= start_date) & (filtered["Date"] <= end_date)
    ]

    return filtered


async def analyze_command(
    i: discord.Interaction,
    file: discord.Attachment,
    start_day: int,
    end_day: int,
) -> None:
    await i.response.defer()
    df = await asyncio.to_thread(
        analyze_excel,
        await file.read(),
        start_day=start_day,
        end_day=end_day,
    )
    if df.empty:
        await i.followup.send("沒有找到符合條件的資料", ephemeral=True)
        return

    df = df.sort_values("Code")
    df = df.reset_index(drop=True)

    row_strs = df.to_string(index=False, header=False).split("\n")
    content = "\n".join(f"{idx + 1}. {row}" for idx, row in enumerate(row_strs))

    txt_file = discord.File(
        fp=io.BytesIO(content.encode("utf-8")),
        filename="results.txt",
    )
    await i.followup.send(file=txt_file)


intents = discord.Intents.default()
bot = commands.Bot(command_prefix=commands.when_mentioned, intents=intents)


@bot.tree.command(name="分析etd", description="獲取距離今天指定天數內的 ETD")
@discord.app_commands.rename(file="檔案", start_day="起始天數", end_day="結束天數")
@discord.app_commands.describe(
    file="要分析的 Excel 檔案",
    start_day="起始天數要取距離今天幾天",
    end_day="結束天數要取距離今天幾天",
)
async def analyze_etd(
    i: discord.Interaction,
    file: discord.Attachment,
    start_day: int,
    end_day: int,
) -> None:
    try:
        await analyze_command(i, file, start_day=start_day, end_day=end_day)
    except Exception as e:
        await i.followup.send(f"發生錯誤: {str(e)}", ephemeral=True)


@bot.tree.command(name="分析特別股", description="獲取距離今天指定天數內的特別股")
@discord.app_commands.rename(file="檔案", start_day="起始天數", end_day="結束天數")
@discord.app_commands.describe(
    file="要分析的 Excel 檔案",
    start_day="起始天數要取距離今天幾天",
    end_day="結束天數要取距離今天幾天",
)
async def analyze_special(
    i: discord.Interaction,
    file: discord.Attachment,
    start_day: int,
    end_day: int,
) -> None:
    try:
        await analyze_command(i, file, start_day=start_day, end_day=end_day)
    except Exception as e:
        await i.response.send_message(f"發生錯誤: {str(e)}", ephemeral=True)


@commands.is_owner()
@bot.command()
async def sync(ctx: commands.Context) -> None:
    await bot.tree.sync()
    await ctx.send("Commands synced successfully!")


if __name__ == "__main__":
    bot.run(os.getenv("TOKEN"))
