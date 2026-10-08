#!/usr/bin/env bash
# Example guest hook for Linux templates (runs as root after the package upgrade and reboot, before sealing).
# Copy, adjust and list it in Templates.<name>.GuestScripts. Files it uses are copied by GuestFiles.
set -euo pipefail

# Internal CA certificates (GuestFiles -> /root/certs)
if compgen -G "/root/certs/*.crt" > /dev/null; then
    if [ -d /usr/local/share/ca-certificates ]; then
        cp /root/certs/*.crt /usr/local/share/ca-certificates/ && update-ca-certificates
    else
        cp /root/certs/*.crt /etc/pki/ca-trust/source/anchors/ && update-ca-trust
    fi
    echo "CA certificates installed"
fi

# Time synchronisation
if systemctl list-unit-files | grep -qE '^chronyd?\.service'; then
    printf 'server ntp1.contoso.local iburst\nserver ntp2.contoso.local iburst\n' > /etc/chrony.d/contoso.conf 2>/dev/null \
        || printf 'server ntp1.contoso.local iburst\nserver ntp2.contoso.local iburst\n' >> /etc/chrony.conf
    systemctl restart chronyd 2>/dev/null || systemctl restart chrony
else
    sed -i 's/^#\?NTP=.*/NTP=ntp1.contoso.local ntp2.contoso.local/' /etc/systemd/timesyncd.conf
    systemctl restart systemd-timesyncd
fi
echo "time sync configured"

# Agents shipped as packages (GuestFiles -> /root/agents), installed once
for pkg in /root/agents/*.deb /root/agents/*.rpm; do
    [ -e "$pkg" ] || continue
    case "$pkg" in
        *.deb) DEBIAN_FRONTEND=noninteractive apt-get install -yq "$pkg" ;;
        *.rpm) dnf -y install "$pkg" 2>/dev/null || yum -y install "$pkg" ;;
    esac
    echo "installed $(basename "$pkg")"
done

# Agents that register on first start must not keep the template's identity
# (example: stop the service and clear its state; adjust to your agents)
# systemctl stop example-agent && rm -rf /var/lib/example-agent/*

rm -rf /root/certs /root/agents
