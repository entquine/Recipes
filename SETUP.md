# What's Cooking with Brian: setup guide

About 30 minutes, on a computer, once. You need no code skills. You only copy and paste.

## What you're setting up

| Piece | Service | Cost | Why this one |
|---|---|---|---|
| Shared ticks and the two logins | **Firebase** (Google) | Free "Spark" plan. No credit card. | Built-in offline support: ticks made with no signal are stored on the phone and sent when signal returns. Free projects stay up when nobody uses them. (Supabase's free projects pause after a week of no use, which is bad for an app you use a few times a week.) |
| The app itself | **GitHub Pages** | Free | You upload files in a web browser, so you don't need to install any tools. The address stays the same. Firebase Hosting would also work, but it needs a command-line tool on your computer. |
| Your recipes | **Google Sheets** | Free | You edit the recipes like a spreadsheet, and both phones get the changes. |

**How you sign in:** each of you has your own email and password, which you create yourself in Firebase. Nobody else can sign up. The database rules let **only your two email addresses** read or change ticks, so anyone else who finds the app address sees only a sign-in screen. You sign in once per phone and stay signed in.

---

## Step 1. Unzip the app

Unzip `recipe-app.zip`. You get a folder named `recipe-app` with these files: `index.html`, `app.js`, `config.js`, `styles.css`, `sw.js`, `manifest.webmanifest`, `recipes.xlsx`, `icons/`, and `source/`.

## Step 2. Create Firebase: the database and logins

1. Go to **https://console.firebase.google.com** and sign in with any Google account.
2. Click **Create a project** (it may say "Get started with a Firebase project"). Name it, for example `dinner-recipes`. When asked about **Google Analytics**, turn it **off**. Click **Create project**.
3. **Add the logins.** In the left menu open **Build → Authentication** and click **Get started**.
   - On the **Sign-in method** tab, choose **Email/Password**, switch on the first toggle (Email/Password), and click **Save**.
   - On the **Users** tab, click **Add user**. Enter your email and a password. Repeat for your wife. Write the passwords down.
   - Optional extra lock: under **Settings → User actions**, untick **Enable create (sign-up)** if you see it. This stops anyone else from creating an account.
4. **Create the database.** Open **Build → Firestore Database** and click **Create database**.
   - If asked for an edition, choose **Standard**.
   - Pick a location near you, for example `nam5 (United States)`. You can't change this later, but any choice works.
   - Choose **Start in production mode**, then click **Create**.
5. **Lock it to the two of you.** In Firestore, open the **Rules** tab. Delete everything there and paste the text below. **Replace the two emails with yours, in lower-case.** Then click **Publish**.

   ```
   rules_version = '2';
   service cloud.firestore {
     match /databases/{database}/documents {
       function isUs() {
         return request.auth != null
           && request.auth.token.email.lower() in ['you@example.com', 'wife@example.com'];
       }
       match /{document=**} {
         allow read, write: if isUs();
       }
     }
   }
   ```
6. **Get the connection details for the app.** Click the **gear icon → Project settings**. Under "Your apps", click the **`</>`** (Web) icon. Name it `dinner`. Leave "Firebase Hosting" **unticked**, then click **Register app**.
   You'll see a block that contains `const firebaseConfig = { apiKey: "...", authDomain: "...", ... };`. Keep this page open, because you need it in step 4.
   (This information is safe to make public. The rules from step 5 are what protect your data.)

## Step 3. Put the app on GitHub Pages

1. Create a free account at **https://github.com**.
2. Click **+ (top right) → New repository**. Name it `dinner`, set it to **Public** (free Pages hosting needs this), and click **Create repository**.
   *Public means anyone could see the app files and the starter recipes.xlsx. Your ticks and login stay private.*
3. On the new repository page, click the **uploading an existing file** link. Open the `recipe-app` folder on your computer, select **everything inside it** (including the `icons` and `source` folders), and drag it onto the page. Wait for the upload to finish, then click **Commit changes**.
4. Open **Settings → Pages**. Under "Build and deployment", set Source to **Deploy from a branch**, set Branch to **main** and **/ (root)**, then click **Save**.
5. After 1–2 minutes, the page shows **"Your site is live at https://YOUR-NAME.github.io/dinner/"**. That is your app's address.

## Step 4. Connect the app to Firebase

1. In your GitHub repository, click **config.js**, then click the **pencil icon** (Edit).
2. Copy the six lines from the Firebase page (apiKey, authDomain, projectId, storageBucket, messagingSenderId, appId) into the `firebase: { ... }` braces. The result should look like this:
   ```js
   window.APP_CONFIG = {
     firebase: {
       apiKey: "AIzaSyB...",
       authDomain: "dinner-recipes-1234.firebaseapp.com",
       projectId: "dinner-recipes-1234",
       storageBucket: "dinner-recipes-1234.firebasestorage.app",
       messagingSenderId: "1234567890",
       appId: "1:1234567890:web:abc123"
     }
   };
   ```
3. Click **Commit changes**. Wait about a minute for GitHub to publish the change.

## Step 5. Install it on both phones

On each phone, while you have signal:
1. Open **Chrome** and go to your app address (`https://YOUR-NAME.github.io/dinner/`).
2. Sign in with that person's email and password.
3. Tap **⋮ (menu) → Add to Home screen → Install** (it may say "Install app").
4. Open the app from the new home-screen icon. It opens full screen.

Test it: tick something on one phone. It should appear on the other within a second or two. The badge at the top right shows **Synced**, **Offline · will sync**, or **Sync error**.

> If an older version still shows after you change something on GitHub, close the app and open it again. Updates arrive on the second open.

## Step 6. Edit recipes in Google Sheets (recommended)

1. Go to **https://drive.google.com**, click **New → File upload**, and choose `recipe_app_data.xlsx`.
2. Open it, then click **File → Save as Google Sheets**. This step is important, because Google can't publish an Excel file. Use the new Google Sheets copy from now on.
3. Click **File → Share → Publish to web**. Choose **Entire Document** and **Microsoft Excel (.xlsx)**, then click **Publish** and **OK**. Copy the link it shows. The link contains `/d/e/`.
4. In the app, open **Settings** and paste the link into **Google Sheet link**, then tap **Save link**. The link is saved for **both** phones.

From then on, edit the Google Sheet as usual. The app checks for changes each time you open it, and you can also tap **Reload recipes** in Settings. Google takes up to about 5 minutes to publish an edit.

*Publishing makes the recipes readable by anyone who has that long, unguessable link. Your ticks are still private.*

**If the app says it couldn't load the sheet:** it keeps showing the last recipes it loaded, so cooking is never blocked. Try a different link instead: in the sheet, click **Share → General access → Anyone with the link (Viewer)**, then paste that normal sheet link into the app.

**Rules for editing the sheet:**
- Keep the tab names **Recipes** and **Shopping** and the column headings as they are.
- **New meal:** add rows with a new value in **Meal**. The leading number (for example `5 Salmon Night`) sets the order in the list and is hidden in the app.
- In the Plan text, references like `(Carrot steps 2-6)`, `(Main step 3)` or `(Side step 1)` turn into colored links that jump to the right step. For this to work, the word before "step" must be **Main**, **Side**, or a word from the dish's name.
- Plan times look like `0:27`, or use `Before` for preheating.
- Each tick is tied to the text of its line. If you edit a line's text, that line's tick resets. Nothing else is affected.

---

### Day to day
- **Clear ticks:** use the button at the bottom of each meal, and of the shopping list. The app asks you to confirm, and it clears the ticks on both phones.
- **Keep screen on:** while a meal is open, the screen stays on. You can turn this off in Settings.
- **No signal:** the app still opens, and your ticks are saved on the phone. They sync by themselves when signal returns.
- **Forgot password:** on the sign-in screen, tap "Forgot password?", or set a new password in Firebase → Authentication → Users.

### For a developer later
The readable code is in `source/`. To rebuild, run `cd source && npm install && npm run build`. That writes `app.js` and `sw.js` into the folder above it.
