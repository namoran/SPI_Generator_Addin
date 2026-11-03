#!/usr/bin/env python3
"""Manifest checks for the SPI Generator add-in.

Parses manifest.xml and manifest.dev.xml and asserts the fields the
Office Store submission depends on: well-formed XML, GUID Id, provider,
permissions, host, source locations, and that every resid referenced in
VersionOverrides is defined in Resources. Run through
tools/validate-manifest.mjs (npm run validate).
"""
import re
import sys
import xml.etree.ElementTree as ET

NS = {
    'o': 'http://schemas.microsoft.com/office/appforoffice/1.1',
    'bt': 'http://schemas.microsoft.com/office/officeappbasictypes/1.0',
    'ov': 'http://schemas.microsoft.com/office/taskpaneappversionoverrides',
}
GUID_RE = re.compile(r'^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$')
PROD_BASE = 'https://namoran.github.io/SPI_Generator_Addin/'
DEV_BASE = 'https://localhost:3000/'

problems = []


def check(path, base):
    try:
        tree = ET.parse(path)
    except ET.ParseError as e:
        problems.append(f'{path}: not well-formed XML: {e}')
        return
    root = tree.getroot()
    if root.tag != '{%s}OfficeApp' % NS['o']:
        problems.append(f'{path}: root is not OfficeApp')
        return

    def text(tag):
        el = root.find(f'o:{tag}', NS)
        return el.text.strip() if el is not None and el.text else None

    def dv(tag):
        el = root.find(f'o:{tag}', NS)
        return el.get('DefaultValue') if el is not None else None

    if not GUID_RE.match(text('Id') or ''):
        problems.append(f'{path}: Id is not a GUID: {text("Id")}')
    if text('ProviderName') != 'Moran Rocks LLC':
        problems.append(f'{path}: ProviderName is {text("ProviderName")!r}')
    if not re.match(r'^\d+\.\d+\.\d+\.\d+$', text('Version') or ''):
        problems.append(f'{path}: Version is {text("Version")!r}')
    if text('Permissions') != 'ReadWriteDocument':
        problems.append(f'{path}: Permissions is {text("Permissions")!r}')
    hosts = [h.get('Name') for h in root.findall('o:Hosts/o:Host', NS)]
    if hosts != ['Workbook']:
        problems.append(f'{path}: Hosts are {hosts}')
    for tag in ('DisplayName', 'Description', 'IconUrl', 'HighResolutionIconUrl', 'SupportUrl'):
        if not dv(tag):
            problems.append(f'{path}: missing {tag} DefaultValue')
    src = root.find('o:DefaultSettings/o:SourceLocation', NS)
    if src is None or not (src.get('DefaultValue') or '').startswith(base):
        problems.append(f'{path}: SourceLocation does not start with {base}')
    for tag in ('IconUrl', 'HighResolutionIconUrl'):
        if not (dv(tag) or '').startswith(base):
            problems.append(f'{path}: {tag} does not start with {base}')

    # Every resid used in VersionOverrides must be defined in Resources,
    # and every resource URL must sit under the same base.
    defined = set()
    for el in root.iter():
        if el.get('id'):
            defined.add(el.get('id'))
    used = set()
    for el in root.iter():
        if el.get('resid'):
            used.add(el.get('resid'))
    missing = used - defined
    if missing:
        problems.append(f'{path}: resids referenced but not defined: {sorted(missing)}')
    resources = root.find('.//ov:Resources', NS)
    for el in resources.iter() if resources is not None else []:
        if el.get('id') and el.get('DefaultValue') is not None:
            v = el.get('DefaultValue') or ''
            if not v.startswith(base) and el.tag in (
                    '{%s}Image' % NS['bt'], '{%s}Url' % NS['bt']):
                problems.append(f'{path}: resource {el.get("id")} URL {v!r} not under {base}')


check('manifest.xml', PROD_BASE)
check('manifest.dev.xml', DEV_BASE)

if problems:
    print('MANIFEST PROBLEMS:')
    for p in problems:
        print(' -', p)
    sys.exit(1)
print('manifest.xml and manifest.dev.xml: well-formed, all checks passed')
