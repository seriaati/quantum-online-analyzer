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
    type: Literal["etd", "special"],
    days: Optional[int] = None,
    start_day: Optional[int] = None,
    end_day: Optional[int] = None,
) -> pd.DataFrame:
    if days is None and (start_day is None or end_day is None):
        raise ValueError("需要提供 'days' 或 'start_day' 和 'end_day' 其中之一")
    if days is not None and (start_day is not None or end_day is not None):
        raise ValueError("不能同時提供 'days' 和 'start_day'/'end_day'")
    if start_day is not None and end_day is not None and start_day > end_day:
        raise ValueError("start_day 不能大於 end_day")

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

    now = pd.Timestamp.now()
    if days is not None:
        start_date = now
        end_date = now + pd.Timedelta(days=days)
    else:  # start_day and end_day are not None
        start_date = now + pd.Timedelta(days=start_day)
        end_date = now + pd.Timedelta(days=end_day)

    filtered = filtered[
        (filtered["Date"] >= start_date) & (filtered["Date"] <= end_date)
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
    days: Optional[int],
    start_day: Optional[int],
    end_day: Optional[int],
    type: Literal["etd", "special"] = "etd",
) -> None:
    await i.response.defer()
    df = await asyncio.to_thread(
        analyze_excel,
        await file.read(),
        type=type,
        days=days,
        start_day=start_day,
        end_day=end_day,
    )
    if df.empty:
        await i.response.send_message("沒有找到符合條件的資料", ephemeral=True)
        return

    df = df.sort_values("Code")
    df = df.reset_index(drop=True)
    paginator = DataFramePaginator(df)
    await i.followup.send(content=paginator.page_content, view=paginator)


intents = discord.Intents.default()
bot = commands.Bot(command_prefix=commands.when_mentioned, intents=intents)


@bot.tree.command(name="分析etd", description="獲取距離今天指定天數內的 ETD")
@discord.app_commands.rename(
    file="檔案", days="天數", start_day="起始天數", end_day="結束天數"
)
@discord.app_commands.describe(
    file="要分析的 Excel 檔案",
    days="要取距離今天的天數",
    start_day="起始天數要取距離今天幾天",
    end_day="結束天數要取距離今天幾天",
)
async def analyze_etd(
    i: discord.Interaction,
    file: discord.Attachment,
    days: Optional[discord.app_commands.Range[int, 365, 1000]] = None,
    start_day: Optional[int] = None,
    end_day: Optional[int] = None,
) -> None:
    try:
        await analyze_command(
            i, file, type="etd", days=days, start_day=start_day, end_day=end_day
        )
    except Exception as e:
        await i.response.send_message(f"發生錯誤: {str(e)}", ephemeral=True)


@bot.tree.command(name="分析特別股", description="獲取距離今天指定天數內的特別股")
@discord.app_commands.rename(
    file="檔案", days="天數", start_day="起始天數", end_day="結束天數"
)
@discord.app_commands.describe(
    file="要分析的 Excel 檔案",
    days="要取距離今天的天數",
    start_day="起始天數要取距離今天幾天",
    end_day="結束天數要取距離今天幾天",
)
async def analyze_special(
    i: discord.Interaction,
    file: discord.Attachment,
    days: Optional[discord.app_commands.Range[int, 365, 1000]] = None,
    start_day: Optional[int] = None,
    end_day: Optional[int] = None,
) -> None:
    try:
        await analyze_command(
            i, file, type="special", days=days, start_day=start_day, end_day=end_day
        )
    except Exception as e:
        await i.response.send_message(f"發生錯誤: {str(e)}", ephemeral=True)


@commands.is_owner()
@bot.command()
async def sync(ctx: commands.Context) -> None:
    await bot.tree.sync()
    await ctx.send("Commands synced successfully!")


if __name__ == "__main__":
    bot.run(os.getenv("TOKEN"))
