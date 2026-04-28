# systemd

`docker-compose@sap.service` starts the Docker Compose stack at boot.

Expected deployment layout on the VM:

```text
/opt/sap/current/deploy
```

`current` is expected to be a symlink to the active release directory. If the VM
uses a plain git checkout instead, update `WorkingDirectory` in the unit before
enabling it.

Install with:

```sh
sudo cp deploy/systemd/docker-compose@sap.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now docker-compose@sap.service
```
