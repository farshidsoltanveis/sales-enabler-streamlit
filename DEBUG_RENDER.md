# Debug: OpenAI Not Working on Render

## Quick Diagnostic

### Step 1: Check Environment Variable Status

Visit this URL on your Render deployment:
```
https://your-app-name.onrender.com/api/health
```

This will show:
- ✅ If `OPENAI_API_KEY` is set
- ✅ Preview of the key (first/last chars for verification)
- ✅ Current environment settings

**Expected Output (if working):**
```json
{
  "status": "ok",
  "openai_api_key": "sk-proj...xyz",
  "openai_configured": true,
  "flask_env": "production",
  "port": "8080"
}
```

**If API key is missing:**
```json
{
  "status": "ok",
  "openai_api_key": "Not Set",
  "openai_configured": false,
  ...
}
```

### Step 2: Check Render Logs

1. Go to Render Dashboard → Your Service → **Logs** tab
2. Look for these messages when processing a file:

**If API key is missing:**
```
ERROR: OPENAI_API_KEY environment variable is not set. Cannot use OpenAI summarization.
⚠️  OpenAI API key not configured: OPENAI_API_KEY environment variable is not set...
   Using offline summary for filename.pdf
```

**If API key is invalid:**
```
⚠️  OpenAI summarization failed for filename.pdf: AuthenticationError: ...
   Falling back to offline summary
```

**If working correctly:**
```
✓ Successfully generated OpenAI summary for filename.pdf
  (OpenAI usage: prompt=123, completion=456, total=579)
```

### Step 3: Verify Environment Variable in Render

1. **Render Dashboard** → Your Service → **Environment** tab
2. **Look for `OPENAI_API_KEY`**
   - Should be listed
   - Should have "Secret" badge/icon
   - Value should be hidden (shows as `•••••`)

3. **If Missing:**
   - Click **"Add Environment Variable"**
   - **Key:** `OPENAI_API_KEY` (exactly, case-sensitive)
   - **Value:** Your API key from https://platform.openai.com/api-keys
   - **Toggle "Secret"** to ON ✅
   - Click **"Save Changes"**

4. **If Exists but Not Working:**
   - Click the variable to edit
   - Verify the value is correct (starts with `sk-`)
   - Make sure there are no extra spaces
   - Click **"Save Changes"**

### Step 4: Redeploy After Changes

**After adding/updating the environment variable:**

1. Go to **"Manual Deploy"** tab
2. Click **"Deploy latest commit"**
3. Wait for deployment (2-5 minutes)
4. Test again

**OR** (if auto-deploy is enabled):
- Just wait for auto-deploy (happens automatically)

## Common Issues & Solutions

### Issue 1: Variable Not Set
**Symptom:** `/api/health` shows `"openai_configured": false`

**Solution:**
1. Add `OPENAI_API_KEY` in Render Environment tab
2. Redeploy
3. Verify with `/api/health` endpoint

### Issue 2: Variable Name Typo
**Symptom:** Variable exists but still not working

**Solution:**
- Must be exactly: `OPENAI_API_KEY` (case-sensitive, no spaces)
- Check for typos: `OPENAI_API_KEY` not `OPENAI_API_KEYS` or `OPEN_AI_API_KEY`

### Issue 3: Invalid API Key
**Symptom:** Logs show `AuthenticationError`

**Solution:**
1. Get a fresh key from: https://platform.openai.com/api-keys
2. Make sure it starts with `sk-`
3. Verify it's active (not revoked)
4. Check you have credits/quota

### Issue 4: Key Not Persisting
**Symptom:** Variable disappears after redeploy

**Solution:**
- Make sure you click **"Save Changes"** after adding
- Don't delete and recreate - just edit existing

### Issue 5: Old Code Deployed
**Symptom:** No error messages in logs

**Solution:**
1. Check latest commit is deployed:
   - Render Dashboard → Your Service → "Events" tab
   - Should show latest commit hash
2. Force redeploy:
   - Manual Deploy → Deploy latest commit

## Step-by-Step Fix

### Complete Fix Process:

1. **Check Current Status**
   ```
   Visit: https://your-app.onrender.com/api/health
   ```

2. **If API Key Missing:**
   - Render Dashboard → Environment → Add `OPENAI_API_KEY`
   - Get key from: https://platform.openai.com/api-keys
   - Mark as Secret ✅
   - Save

3. **Redeploy**
   - Manual Deploy → Deploy latest commit
   - Wait 2-5 minutes

4. **Verify**
   - Check `/api/health` again
   - Should show `"openai_configured": true`

5. **Test**
   - Upload a test file
   - Check logs for success message
   - Verify brief file has `_brief.md` (not `_brief_offline.md`)

## Testing Locally First

Before deploying, test locally:

```bash
# Test without API key (should fail gracefully)
unset OPENAI_API_KEY
python app.py
# Visit http://localhost:5001/api/health
# Should show "openai_configured": false

# Test with API key (should work)
export OPENAI_API_KEY="sk-your-key-here"
python app.py
# Visit http://localhost:5001/api/health
# Should show "openai_configured": true
# Upload a file - should use OpenAI
```

## Verification Checklist

- [ ] `/api/health` shows `"openai_configured": true`
- [ ] API key preview shows correct format (`sk-...`)
- [ ] Render Environment tab shows `OPENAI_API_KEY` with Secret badge
- [ ] Service has been redeployed after adding variable
- [ ] Logs show "Successfully generated OpenAI summary" (not offline)
- [ ] Brief files have `_brief.md` suffix (not `_brief_offline.md`)
- [ ] Brief content shows `### [LLM: OpenAI]` (not `### [LLM: Offline]`)

## Still Not Working?

1. **Share the `/api/health` response** - This will show what's configured
2. **Share relevant log lines** - Copy error messages from Render logs
3. **Verify API key works** - Test it locally with the same key
4. **Check Render status** - Sometimes Render has service issues

## Quick Test Command

You can test if your API key works from command line:

```bash
export OPENAI_API_KEY="sk-your-key-here"
python -c "from openai import OpenAI; client = OpenAI(); print('✅ API key works!'); print(list(client.models.list())[0].id)"
```

If this fails, the API key is invalid.

---

**The `/api/health` endpoint is your best friend for debugging!** 🚀

