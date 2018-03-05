## get-site-members

This is a web part that uses react table to display users from a site group.
The web part properties display available groups, selecting a group will update the state and render the users.

The web part works, but lacks proper strings for languages etc..

NB! The groups displayed should have the "allow everyone to view memberships in this group" option set unless you want to display a 403.

### Demo

![get group members2](https://user-images.githubusercontent.com/20144749/36982526-570cece2-2090-11e8-9cf5-e74192950fd1.gif)

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
