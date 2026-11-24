# NTA Onboarding

## Summary

SPFx React web parts for an onboarding page:

- **OnboardTopBar** displays the NTA brand, employee identity, onboarding day, and task progress.
- **OnboardingHero** displays a full-width welcome section with the journey banner, onboarding summary, story timeline, and journey-map link.

## Run locally

- Use a supported Node.js version for SPFx 1.24.2.
- Run `npm install`.
- Run `npm start`.
- Open the SharePoint workbench and add **OnboardTopBar** or **OnboardingHero**.

## Configure the web part

Use each web part's property pane to configure its text and progress values. If the employee name is left blank, the web part uses the signed-in SharePoint user's display name.

Both web parts use right-to-left layout for Hebrew content and adapt to narrow screens. **OnboardingHero** supports full bleed; place it in a full-width section on a Communication site to extend the background across the page.

## Build and package

Run `npm run build` to validate and package the SPFx solution.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.24.2-green.svg)
