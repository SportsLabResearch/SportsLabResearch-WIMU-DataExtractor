# -*- coding: utf-8 -*-

"""
Gestión de sensores WIMU.
"""

from __future__ import annotations

from typing import Dict, List

from .constants import QUL_ALIAS_SENSOR_CODE, QUL_ORDER
from .utils import safe_text


def get_sensor(packet_type: int, raw_code: int,
               sensors: Dict[int, Dict[str, object]]) -> Dict[str, object]:
    """
    Devuelve la definición del sensor asociada a un paquete.
    """

    sensor_code = QUL_ALIAS_SENSOR_CODE.get(raw_code, raw_code)

    sensor = sensors.get(sensor_code)

    if sensor:
        return sensor

    return {
        "sensor_code": sensor_code,
        "name": f"PACKET_{packet_type}_{raw_code}",
        "type": "",
        "enabled": "",
        "visible": "",
        "magnitude": "",
        "unit": "",
        "channels": [],
    }


def channel_names(sensor: Dict[str, object],
                  value_count: int) -> List[str]:
    """
    Devuelve los nombres de los canales de un sensor.
    """

    names: List[str] = []

    raw_channels = sensor.get("channels", [])

    if isinstance(raw_channels, list):
        for ch in raw_channels:
            if isinstance(ch, dict):
                names.append(
                    safe_text(ch.get("name")) or
                    f"valor_{len(names)+1}"
                )

    while len(names) < value_count:
        names.append(f"valor_{len(names)+1}")

    return names[:value_count]


def packet_sort_key(info: Dict[str, object]):
    """
    Clave de orden para presentar los sensores.
    """

    packet_type = int(info["packet_type"])
    raw_code = int(info["raw_code"])

    return (
        QUL_ORDER.get((packet_type, raw_code), 999),
        packet_type,
        raw_code,
        int(info["size"]),
    )