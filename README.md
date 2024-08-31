Fantasy Football Cheat Sheet
============================

I put together a fantasy football cheat-sheet every year in
order to be able to see value more easily when drafting.
This script fetches tiers data for 0.5 point PPR from
[Fantasy Pros here](https://www.fantasypros.com/nfl/rankings/half-point-ppr-cheatsheets.php).

### TL;DR

1. Run `bundle install` (you'll need ruby)
2. Run `./cheat.rb`

![Running the script](https://raw.githubusercontent.com/unRARed/cheat-sheet/main/example-run.png)

Now, crack a beer and draft. =)

Ranking data is saved to `./tiers.json`. Once scraped, the script
will use this data unless the file is deleted.

### Options

- `./cheat.rb --fresh` force a fresh scraping of the data.
- `./cheat.rb --idp` ensure IDP players are included in the
  scraping step.
- `./cheat.rb --concerns` include a list of players with
  concerns going into the season.

![Example output](https://raw.githubusercontent.com/unRARed/cheat-sheet/main/cheat-sheet-snip.png)

