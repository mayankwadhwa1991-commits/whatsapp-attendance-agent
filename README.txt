WHATSAPP ATTENDANCE AGENT
=========================

1. Deploy this folder as a Web Service on Render.
2. Start command:
   uvicorn main:app --host 0.0.0.0 --port 10000

3. Environment variables to set on Render:
   VERIFY_TOKEN = my_verify_token
   WHATSAPP_TOKEN = your_whatsapp_token
   (Use the token from WhatsApp Cloud API)
   (Replace YOUR_PHONE_NUMBER_ID in main.py with your real phone number ID)

4. Webhook URL on Render:
   https://your-render-app-url/webhook

5. In WhatsApp Dev Console:
   - Go to Configuration → Webhooks
   - Paste the Render webhook URL
   - Set VERIFY_TOKEN = my_verify_token
   - Click Verify and Save

6. Send a real attendance message from WhatsApp to test.