# Fix: OpenAI Not Working on Render

## Problem
Analytics are generated successfully with OpenAI on localhost, but Render is using offline mode.

## Root Cause
The `OPENAI_API_KEY` environment variable is either:
1. **Not set** in Render dashboard
2. **Set incorrectly** (typo, wrong value)
3. **Not accessible** to the application

## Solution

### Step 1: Verify Environment Variable in Render

1. **Go to Render Dashboard**
   - Visit: https://dashboard.render.com
   - Click on your service: `sales-enabler`

2. **Check Environment Variables**
   - Click on **"Environment"** tab
   - Look for `OPENAI_API_KEY`
   - Verify it exists and has a value

3. **If Missing, Add It:**
   - Click **"Add Environment Variable"**
   - **Key:** `OPENAI_API_KEY`
   - **Value:** Your OpenAI API key (starts with `sk-`)
   - **Toggle "Secret" to ON** ✅
   - Click **"Save Changes"**

### Step 2: Verify API Key Format

Your OpenAI API key should:
- Start with `sk-`
- Be about 50+ characters long
- Be from: https://platform.openai.com/api-keys

**Example format:**
```
sk-proj-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

### Step 3: Redeploy After Adding Variable

After adding/updating the environment variable:
1. Go to **"Manual Deploy"** tab
2. Click **"Deploy latest commit"**
3. Wait for deployment to complete
4. Check logs to verify it's working

### Step 4: Check Logs for Errors

1. Go to **"Logs"** tab in Render
2. Look for these messages:

**If API key is missing:**
```
ERROR: OPENAI_API_KEY environment variable is not set. Cannot use OpenAI summarization.
⚠️  OpenAI API key not configured: ...
   Using offline summary for ...
```

**If API key is invalid:**
```
⚠️  OpenAI summarization failed for ...: AuthenticationError: ...
   Falling back to offline summary
```

**If working correctly:**
```
✓ Successfully generated OpenAI summary for ...
  (OpenAI usage: prompt=..., completion=..., total=...)
```

## Updated Code

I've updated the code to:
1. ✅ **Check for API key** before attempting OpenAI calls
2. ✅ **Log clear error messages** when API key is missing
3. ✅ **Distinguish between** missing key vs. other errors
4. ✅ **Explicitly pass API key** to OpenAI client

## Testing

### Test Locally First
```bash
# Without API key (should show error)
unset OPENAI_API_KEY
python app.py
# Upload a file and check logs

# With API key (should work)
export OPENAI_API_KEY="sk-your-key-here"
python app.py
# Upload a file - should use OpenAI
```

### Test on Render
1. Set `OPENAI_API_KEY` in Render dashboard
2. Redeploy
3. Upload a test file
4. Check logs for success message

## Common Issues

### Issue 1: Variable Not Saved
- Make sure to click **"Save Changes"** after adding variable
- Refresh the page and verify it's still there

### Issue 2: Wrong Variable Name
- Must be exactly: `OPENAI_API_KEY` (case-sensitive)
- No spaces, no typos

### Issue 3: Invalid API Key
- Get a fresh key from: https://platform.openai.com/api-keys
- Make sure it's active and has credits

### Issue 4: Key Not Visible
- If marked as "Secret", it won't show the value (this is normal)
- Just verify the key name exists

## Verification Checklist

- [ ] `OPENAI_API_KEY` exists in Render Environment tab
- [ ] Variable is marked as "Secret"
- [ ] Value starts with `sk-`
- [ ] Service has been redeployed after adding variable
- [ ] Logs show "Successfully generated OpenAI summary" (not offline)
- [ ] Brief files have `_brief.md` suffix (not `_brief_offline.md`)

## Still Not Working?

1. **Check Render Logs:**
   - Look for error messages
   - Copy the exact error and share it

2. **Verify API Key:**
   - Test the key locally: `export OPENAI_API_KEY="sk-..." && python -c "from openai import OpenAI; print(OpenAI().models.list())"`
   - If this fails, the key is invalid

3. **Check Billing:**
   - Go to: https://platform.openai.com/account/billing
   - Ensure you have credits/quota available

4. **Contact Support:**
   - Render: support@render.com
   - OpenAI: help.openai.com

---

**After fixing, the app should use OpenAI instead of offline mode!** 🚀

