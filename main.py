import asyncio
import io
import os
from typing import Literal
import discord
from discord.ext import commands

from dotenv import load_dotenv
import pandas as pd

load_dotenv()


def analyze_excel(
    file: bytes, days: int, *, type: Literal["etd", "special"]
) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file))

    filtered = df[["tablescraper-selected-row 3", "tablescraper-selected-row 10"]]

    filtered.columns = ["Code", "Date"]
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

    filtered = filtered[filtered["Date"] >= pd.Timestamp.now()]
    filtered = filtered[
        filtered["Date"] - pd.Timestamp.now() <= pd.Timedelta(days=days)
    ]
    return filtered


class DataFramePaginator(discord.ui.View):
    def __init__(self, df: pd.DataFrame):
        super().__init__(timeout=None)
        self.df = df
        self.page = 0
        self.page_size = 30

        row_strs = df.to_string(index=False, header=False).split("\n")
        row_strs = [f"{i + 1}. {row}" for i, row in enumerate(row_strs)]
        self.pages = [
            row_strs[i : i + self.page_size]
            for i in range(0, len(row_strs), self.page_size)
        ]

    @property
    def page_content(self) -> str:
        return "\n".join(self.pages[self.page])

    async def _update_message(self, i: discord.Interaction):
        await i.response.edit_message(content=self.page_content, view=self)

    @discord.ui.button(label="上一頁", style=discord.ButtonStyle.primary)
    async def previous_page(self, i: discord.Interaction, _: discord.ui.Button):
        self.page -= 1
        if self.page < 0:
            self.page = len(self.pages) - 1
        await self._update_message(i)

    @discord.ui.button(label="下一頁", style=discord.ButtonStyle.primary)
    async def next_page(self, i: discord.Interaction, _: discord.ui.Button):
        self.page += 1
        if self.page >= len(self.pages):
            self.page = 0
        await self._update_message(i)


async def analyze_command(
    i: discord.Interaction,
    file: discord.Attachment,
    days: int = 500,
    type: Literal["etd", "special"] = "etd",
) -> None:
    await i.response.defer()
    df = await asyncio.to_thread(analyze_excel, await file.read(), days, type=type)
    if df.empty:
        await i.response.send_message("沒有找到符合條件的資料。", ephemeral=True)
        return

    df = df.sort_values("Code")
    df = df.reset_index(drop=True)
    paginator = DataFramePaginator(df)
    await i.followup.send(content=paginator.page_content, view=paginator)


intents = discord.Intents.default()
bot = commands.Bot(command_prefix=commands.when_mentioned, intents=intents)


@bot.tree.command(name="分析etd", description="獲取距離今天指定天數內的 ETD")
@discord.app_commands.rename(file="檔案", days="天數")
@discord.app_commands.describe(file="要分析的 Excel 檔案", days="要取距離今天的天數")
async def analyze_etd(
    i: discord.Interaction,
    file: discord.Attachment,
    days: discord.app_commands.Range[int, 365, 1000] = 500,
) -> None:
    try:
        await analyze_command(i, file, days, type="etd")
    except Exception as e:
        await i.response.send_message(f"發生錯誤: {str(e)}", ephemeral=True)


@bot.tree.command(name="分析特別股", description="獲取距離今天指定天數內的特別股")
@discord.app_commands.rename(file="檔案", days="天數")
@discord.app_commands.describe(file="要分析的 Excel 檔案", days="要取距離今天的天數")
async def analyze_special(
    i: discord.Interaction,
    file: discord.Attachment,
    days: discord.app_commands.Range[int, 365, 1000] = 500,
) -> None:
    try:
        await analyze_command(i, file, days, type="special")
    except Exception as e:
        await i.response.send_message(f"發生錯誤: {str(e)}", ephemeral=True)


@commands.is_owner()
@bot.command()
async def sync(ctx: commands.Context) -> None:
    await bot.tree.sync()
    await ctx.send("Commands synced successfully!")


if __name__ == "__main__":
    bot.run(os.getenv("TOKEN"))
