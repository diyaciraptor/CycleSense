# MTrack Android With Capacitor

This folder is ready to be wrapped as an Android app with Capacitor.

## 1. Install prerequisites

Install these on your computer:
- Node.js LTS
- Android Studio
- Java 17 if Android Studio does not install it for you

Then open a terminal in this folder:
`D:\Period Tracker`

## 2. Install Capacitor

Run:

```powershell
npm install
```

## 3. Create the Android project

Run:

```powershell
npx cap add android
```

This creates an `android` folder.

## 4. Sync the web app into Android

Run:

```powershell
npx cap sync android
```

Any time you change `index.html`, `app.js`, `style.css`, `config.js`, or other web files, run:

```powershell
npx cap copy android
```

## 5. Open in Android Studio

Run:

```powershell
npx cap open android
```

Then in Android Studio:
- wait for Gradle sync
- choose an emulator or a connected Android phone
- click Run

## 6. Supabase note

Your `config.js` values are bundled into the Android app, so only use the client-safe Supabase key.
Never use the service role key.

## 7. Recommended test flow

1. Open the Android app.
2. Create an account or sign in.
3. Save a cycle.
4. Confirm the cycle appears in Supabase `cycle_entries`.
5. Make another change and run `npx cap copy android` before rebuilding.

## 8. If you deploy updates later

For web-only changes:

```powershell
npx cap copy android
```

For plugin or config changes:

```powershell
npx cap sync android
```
