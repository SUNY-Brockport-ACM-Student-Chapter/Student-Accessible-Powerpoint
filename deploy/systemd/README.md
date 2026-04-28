# systemd

`docker-compose@sap.service` starts the Docker Compose stack at boot.

Expected deployment layout on the VM:

```text
/opt/sap/current/deploy
```

Install with:

```sh
sudo cp deploy/systemd/docker-compose@sap.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now docker-compose@sap.service
```
