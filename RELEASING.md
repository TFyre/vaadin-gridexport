# Releasing

Releases are published to Maven Central through Sonatype's Central Portal using
`org.sonatype.central:central-publishing-maven-plugin`. The `release` profile builds, signs and
uploads a deployment, but does **not** publish it automatically (`autoPublish=false`) — a
maintainer reviews the validated deployment in the Portal and publishes it manually.

## Prerequisites

1. **GPG key** — the signing key must be published to a public keyserver (for example
   `keys.openpgp.org`) and its passphrase must be available to `maven-gpg-plugin`, either through
   `gpg-agent` or `-Dgpg.passphrase=...`. The key must meet Central's signing requirements.
2. **Central Portal token** — generate a user token at <https://central.sonatype.com/usertoken>
   and add it to `~/.m2/settings.xml` under the server id `central`:

   ```xml
   <settings>
     <servers>
       <server>
         <id>central</id>
         <username>TOKEN_USERNAME</username>
         <password>TOKEN_PASSWORD</password>
       </server>
     </servers>
   </settings>
   ```

3. **Verified namespace** — the `com.tfyre` namespace must be verified in the Portal
   (account → Namespaces).
4. **Clean working tree** on `master` at the commit to release, with the POM version already set
   to the release version (no `-SNAPSHOT`).

> Never commit tokens or GPG secrets. Keep `settings.xml` outside the repository.

## Release steps

1. Update `master`, then create and push an annotated tag whose name matches the POM version:

   ```bash
   git checkout master
   git pull
   git tag -a v3.0.0 -m "Release 3.0.0"
   git push origin v3.0.0
   ```

2. Build, sign and upload the deployment from the repository root:

   ```bash
   mvn -Prelease deploy
   ```

   The plugin uploads the bundle and waits until it is validated. Because `autoPublish=false`,
   nothing is published yet; the build output links to the deployment in the Portal.

3. Publish manually: open <https://central.sonatype.com/publishing/deployments>, select the
   validated deployment and click **Publish**. Central then syncs the artifacts to Maven Central
   (usually within 15–30 minutes).

## Dry run

To verify the release build without signing or Portal credentials:

```bash
mvn -Prelease -DskipTests -Dgpg.skip=true verify
```

This runs compilation, sources and Javadoc generation; it does not contact the Portal.
