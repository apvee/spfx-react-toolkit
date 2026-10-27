/*
 * Mapping data: CLDR48 territory001, pinned at
 * acd6d88ae493633240e19a87a721076a8a75c310. Numeric candidate identities:
 * PnP Core b6ff7b30dc8a1dfa8202bca50b75a08e3d3d6c43; Microsoft corrections.
 * Audit, source snapshots and exclusions: tests/fixtures/spfx-time-zone-mapping.json.
 *
 * UNICODE LICENSE V3
 *
 * COPYRIGHT AND PERMISSION NOTICE
 *
 * Copyright © 2004-2025 Unicode, Inc.
 *
 * NOTICE TO USER: Carefully read the following legal agreement. BY
 * DOWNLOADING, INSTALLING, COPYING OR OTHERWISE USING DATA FILES, AND/OR
 * SOFTWARE, YOU UNEQUIVOCALLY ACCEPT, AND AGREE TO BE BOUND BY, ALL OF THE
 * TERMS AND CONDITIONS OF THIS AGREEMENT. IF YOU DO NOT AGREE, DO NOT
 * DOWNLOAD, INSTALL, COPY, DISTRIBUTE OR USE THE DATA FILES OR SOFTWARE.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a
 * copy of data files and any associated documentation (the "Data Files") or
 * software and any associated documentation (the "Software") to deal in the
 * Data Files or Software without restriction, including without limitation
 * the rights to use, copy, modify, merge, publish, distribute, and/or sell
 * copies of the Data Files or Software, and to permit persons to whom the
 * Data Files or Software are furnished to do so, provided that either (a)
 * this copyright and permission notice appear with all copies of the Data
 * Files or Software, or (b) this copyright and permission notice appear in
 * associated Documentation.
 *
 * THE DATA FILES AND SOFTWARE ARE PROVIDED "AS IS", WITHOUT WARRANTY OF ANY
 * KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT OF
 * THIRD PARTY RIGHTS.
 *
 * IN NO EVENT SHALL THE COPYRIGHT HOLDER OR HOLDERS INCLUDED IN THIS NOTICE
 * BE LIABLE FOR ANY CLAIM, OR ANY SPECIAL INDIRECT OR CONSEQUENTIAL DAMAGES,
 * OR ANY DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS,
 * WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION,
 * ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THE DATA
 * FILES OR SOFTWARE.
 *
 * Except as contained in this notice, the name of a copyright holder shall
 * not be used in advertising or otherwise to promote the sale, use or other
 * dealings in these Data Files or Software without prior written
 * authorization of the copyright holder.
 *
 * SPDX-License-Identifier: Unicode-3.0
 *
 * MIT License
 *
 * Copyright (c) 2020 PnP
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *
 */

// Standard group representatives only; unknown/customized groups remain unresolved.
const spfxIanaTimeZones: Readonly<Partial<Record<number, string>>> = {
  2: 'Europe/London',
  3: 'Europe/Paris',
  4: 'Europe/Berlin',
  5: 'Europe/Bucharest',
  6: 'Europe/Budapest',
  8: 'America/Sao_Paulo',
  9: 'America/Halifax',
  10: 'America/New_York',
  11: 'America/Chicago',
  12: 'America/Denver',
  13: 'America/Los_Angeles',
  14: 'America/Anchorage',
  15: 'Pacific/Honolulu',
  17: 'Pacific/Auckland',
  18: 'Australia/Brisbane',
  19: 'Australia/Adelaide',
  20: 'Asia/Tokyo',
  21: 'Asia/Singapore',
  22: 'Asia/Bangkok',
  23: 'Asia/Calcutta',
  24: 'Asia/Dubai',
  25: 'Asia/Tehran',
  26: 'Asia/Baghdad',
  27: 'Asia/Jerusalem',
  28: 'America/St_Johns',
  29: 'Atlantic/Azores',
  31: 'Atlantic/Reykjavik',
  33: 'America/La_Paz',
  34: 'America/Indianapolis',
  35: 'America/Bogota',
  36: 'America/Regina',
  37: 'America/Mexico_City',
  38: 'America/Phoenix',
  39: 'Etc/GMT+12',
  40: 'Pacific/Fiji',
  42: 'Australia/Hobart',
  43: 'Pacific/Port_Moresby',
  44: 'Australia/Darwin',
  45: 'Asia/Shanghai',
  46: 'Asia/Novosibirsk',
  47: 'Asia/Tashkent',
  48: 'Asia/Kabul',
  49: 'Africa/Cairo',
  50: 'Africa/Johannesburg',
  51: 'Europe/Moscow',
  53: 'Atlantic/Cape_Verde',
  54: 'Asia/Baku',
  55: 'America/Guatemala',
  56: 'Africa/Nairobi',
  57: 'Europe/Warsaw',
  58: 'Asia/Yekaterinburg',
  59: 'Europe/Kiev',
  60: 'America/Godthab',
  61: 'Asia/Rangoon',
  62: 'Asia/Katmandu',
  63: 'Asia/Irkutsk',
  64: 'Asia/Krasnoyarsk',
  65: 'America/Santiago',
  66: 'Asia/Colombo',
  67: 'Pacific/Tongatapu',
  68: 'Asia/Vladivostok',
  69: 'Africa/Lagos',
  70: 'Asia/Yakutsk',
  71: 'Asia/Bishkek',
  72: 'Asia/Seoul',
  73: 'Australia/Perth',
  74: 'Asia/Riyadh',
  75: 'Asia/Taipei',
  76: 'Australia/Sydney',
  77: 'America/Mazatlan',
  78: 'America/Tijuana',
  79: 'Asia/Amman',
  80: 'Asia/Beirut',
  82: 'Asia/Tbilisi',
  83: 'Africa/Windhoek',
  84: 'Asia/Yerevan',
  85: 'America/Buenos_Aires',
  86: 'Africa/Casablanca',
  87: 'Asia/Karachi',
  88: 'America/Caracas',
  89: 'Indian/Mauritius',
  90: 'America/Montevideo',
  91: 'America/Asuncion',
  93: 'Etc/UTC',
  94: 'Asia/Ulaanbaatar',
  95: 'Etc/GMT+11',
  96: 'Etc/GMT+2',
  97: 'Etc/GMT-12',
  98: 'Asia/Damascus',
  99: 'Asia/Magadan',
  100: 'Europe/Kaliningrad',
  101: 'Europe/Istanbul',
  102: 'Asia/Dhaka',
  103: 'America/Bahia',
  104: 'Europe/Chisinau',
  106: 'Europe/Samara',
  107: 'Asia/Srednekolymsk',
  108: 'Asia/Kamchatka',
  109: 'Europe/Minsk',
  110: 'Europe/Astrakhan',
  111: 'Asia/Barnaul',
  112: 'Asia/Tomsk',
  114: 'Asia/Sakhalin',
  115: 'Asia/Omsk',
};

/** Resolves only audited standard SharePoint numeric time-zone IDs. */
export function resolveSPFxIanaTimeZone(id: number | undefined): string | undefined {
  return typeof id === 'number' && Number.isInteger(id) ? spfxIanaTimeZones[id] : undefined;
}
