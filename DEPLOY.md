# Multi-Tenant Banneker News Tracker — Deploy Guide

End state: One Streamlit app where any Banneker coworker can self-serve their own personalized news tracker in ~4 min, with an optional weekly/daily email delivery. All free. Hosted on your accounts.

## One-time setup (~25 min, only you do this)

### 1. Create the Gist that holds all tracker configs (~3 min)

1. Go to https://gist.github.com
2. **Filename:** `trackers.json`
3. **Content:** `{}` (literally just an empty JSON object)
4. Click **Create secret gist** (or public — doesn't matter much, configs aren't sensitive)
5. From the URL, copy the long hex Gist ID — looks like `https://gist.github.com/natefoster32/<GIST_ID_HERE>`

### 2. Create a GitHub Personal Access Token (~2 min)

1. Go to https://github.com/settings/tokens?type=beta (fine-grained tokens)
2. Click **Generate new token**
3. Name: `news-tracker-gist`
4. Expiration: 1 year (set a calendar reminder to rotate)
5. **Repository access:** No specific repos needed (Gists are user-scope)
6. **Account permissions:** `Gists` → **Read and write**
7. Generate, copy the token (starts with `github_pat_...`)

If fine-grained tokens give you trouble, classic tokens with the `gist` scope also work.

### 3. Sign up for Resend for email delivery (~5 min)

1. Go to https://resend.com — sign up (free tier: 3,000 emails/month, no credit card)
2. **For testing:** Use the built-in `onboarding@resend.dev` as the From address — works immediately, no DNS setup
3. **For production:** Add your domain and verify DNS records (3 TXT/MX records). Recommended sender: `brief@banneker.com` or similar
4. **API Keys** → Create API Key → name `news-tracker` → copy (starts with `re_...`)

### 4. Push the new code to your GitHub repo (~3 min)

You're on the `multi-tenant` branch. From the repo:

```powershell
git push -u origin multi-tenant
```

Then on GitHub.com, open a PR from `multi-tenant` to `main` and merge it. (Or `git checkout main; git merge multi-tenant; git push`.)

### 5. Configure Streamlit Cloud Secrets (~3 min)

1. Go to https://share.streamlit.io
2. Open your app **ot-cyber-brief** (or whatever you named it)
3. **Settings** → **Secrets**
4. Paste this, filling in your values:

   ```toml
   GITHUB_GIST_ID = "your_gist_id_here"
   GITHUB_PAT = "github_pat_..."
   RESEND_API_KEY = "re_..."
   RESEND_FROM = "Banneker Brief <onboarding@resend.dev>"
   ```

5. Save. Streamlit will redeploy.

### 6. (Optional but recommended) Rename the app for clarity (~2 min)

The current URL `ot-cyber-brief.streamlit.app` is fine but the app is now multi-tenant. To rename:

1. Streamlit Cloud → your app → Settings → **Custom subdomain**
2. Change to `banneker-news` → URL becomes `banneker-news.streamlit.app`
3. Update Bryan's bookmark: the OT cyber tracker's permanent URL becomes `banneker-news.streamlit.app/?id=<id_you_give_it>`

### 7. Configure GitHub Actions for email cron (~5 min)

This step is only needed if you want the email feature live.

1. In your `ot-cyber-brief` repo on GitHub: **Settings** → **Secrets and variables** → **Actions**
2. Add these repository secrets:

   | Name | Value |
   |---|---|
   | `GIST_ID` | your Gist ID |
   | `GIST_PAT` | your GitHub PAT |
   | `RESEND_API_KEY` | your Resend key |
   | `RESEND_FROM` | `Banneker Brief <onboarding@resend.dev>` (or your verified domain) |
   | `APP_BASE_URL` | `https://banneker-news.streamlit.app` |

3. The workflow at `.github/workflows/send_emails.yml` is already set to run daily at 11:00 UTC. To test immediately: **Actions** tab → **Send daily news briefs** → **Run workflow**

### 8. Migrate Bryan's tracker (~2 min)

The previous version had Bryan's OT cyber queries hardcoded. To recreate it on the new multi-tenant version:

1. Visit your live app → click **Create a new tracker**
2. Pick the **Cybersecurity portco** template (it's pre-loaded with the OT cyber queries from before)
3. Name it "Industrial Defender Weekly Brief" (or whatever)
4. Subtitle: "Banneker Partners · Industrial Defender Market Intel"
5. Save
6. Copy the resulting URL → send to Bryan

His new permanent URL: `https://banneker-news.streamlit.app/?id=industrial-defender-weekly-brief`

## Sharing with coworkers

Just send this Slack message:

> Hey team — built a tool. If you want a personalized weekly news brief on your patch, your deal, or your portco, go here: **banneker-news.streamlit.app**. Pick "Create a new tracker," pick a portco template (cyber / healthcare / fintech / climate / etc.), edit the queries to match your actual focus, save. You get a bookmarkable URL. Click "Email me" inside your tracker if you want it sent weekly/daily to your inbox. Free, no signups. DM me if you want help tuning queries.

## Troubleshooting

- **"GITHUB_GIST_ID is not set" error:** Streamlit Secrets aren't saved or the app hasn't redeployed. Settings → Secrets → re-save → wait 30 seconds.
- **Gist API 404:** Wrong Gist ID, or your PAT doesn't have gist scope.
- **Resend 422 error:** From address isn't verified. Either use `onboarding@resend.dev` for now or finish domain verification.
- **No emails sent in cron run:** Check the Actions log. Common cause: no tracker has an active email subscription yet, OR the frequency doesn't match today's weekday.
- **Coworker sees "Tracker not found":** The Gist write failed silently. Check the Gist on github.com/gists to confirm their config is there.

## Cost summary

All free, indefinitely:
- Streamlit Community Cloud (unlimited apps, low-traffic friendly)
- GitHub Gists (unlimited)
- GitHub Actions (~5 min/day = ~150 min/month, well within 2,000 min free tier)
- Resend (3,000 emails/month free)
- GitHub repo (free for public)

If you ever outgrow free tiers, the natural upgrade is Streamlit Teams ($20/mo) for private apps and Resend Pro ($20/mo for 50k emails). Even at that scale you're at ~$40/mo to run the whole thing.
