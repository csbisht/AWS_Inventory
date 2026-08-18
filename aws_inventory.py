#!/usr/bin/env python3
"""
AWS Resource Inventory
Pulls EC2, Security Groups, VPC, IGW, NAT GW, Transit GW, and S3 into CSVs.

Usage:
    python aws_inventory.py --region us-east-1 [--profile myprofile] [--outdir ./inventory]
    python aws_inventory.py --all-regions [--profile myprofile]

Requires: boto3, valid AWS credentials (env vars / profile / instance role)
"""

import argparse
import csv
import os
import sys
from datetime import datetime

import boto3
from botocore.exceptions import ClientError, NoCredentialsError


def tags_to_str(tags):
    if not tags:
        return ""
    return "; ".join(f"{t['Key']}={t['Value']}" for t in tags)


def name_from_tags(tags):
    if not tags:
        return ""
    for t in tags:
        if t["Key"] == "Name":
            return t["Value"]
    return ""


def sanitize_filename(name):
    return "".join(c if c.isalnum() or c in "-_." else "_" for c in name)


def write_resource_csv(outdir, subfolder, resource_id, resource_name, fields_dict):
    folder = os.path.join(outdir, subfolder)
    os.makedirs(folder, exist_ok=True)
    label = f"{resource_id}_{resource_name}" if resource_name else resource_id
    filename = sanitize_filename(label)[:150] + ".csv"
    path = os.path.join(folder, filename)
    with open(path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Field", "Value"])
        for k, v in fields_dict.items():
            writer.writerow([k, v])
    return path


def write_resource_group(outdir, subfolder, rows, id_field, name_field=None):
    count = 0
    for row in rows:
        write_resource_csv(outdir, subfolder, row[id_field],
                            row.get(name_field, "") if name_field else "", row)
        count += 1
    print(f"  -> {subfolder}/  ({count} files)")


# ---------------- EC2 Instances ----------------
def collect_ec2(ec2, region):
    rows = []
    paginator = ec2.get_paginator("describe_instances")
    for page in paginator.paginate():
        for res in page["Reservations"]:
            for i in res["Instances"]:
                rows.append({
                    "Region": region,
                    "InstanceId": i["InstanceId"],
                    "Name": name_from_tags(i.get("Tags")),
                    "State": i["State"]["Name"],
                    "InstanceType": i["InstanceType"],
                    "PrivateIp": i.get("PrivateIpAddress", ""),
                    "PublicIp": i.get("PublicIpAddress", ""),
                    "VpcId": i.get("VpcId", ""),
                    "SubnetId": i.get("SubnetId", ""),
                    "AZ": i.get("Placement", {}).get("AvailabilityZone", ""),
                    "LaunchTime": str(i.get("LaunchTime", "")),
                    "SecurityGroups": "; ".join(sg["GroupId"] for sg in i.get("SecurityGroups", [])),
                    "Tags": tags_to_str(i.get("Tags")),
                })
    return rows


# ---------------- Security Groups (one row per SG, rules consolidated) ----------------
def format_rule(rule):
    proto = rule.get("IpProtocol", "-1")
    proto = "All" if proto == "-1" else proto
    from_p = rule.get("FromPort", "All")
    to_p = rule.get("ToPort", "All")
    port = "All" if from_p in ("All", None) else (f"{from_p}" if from_p == to_p else f"{from_p}-{to_p}")
    sources = [r["CidrIp"] for r in rule.get("IpRanges", [])]
    sources += [r["CidrIpv6"] for r in rule.get("Ipv6Ranges", [])]
    sources += [r["GroupId"] for r in rule.get("UserIdGroupPairs", [])]
    sources += [r["PrefixListId"] for r in rule.get("PrefixListIds", [])]
    src = ",".join(sources) if sources else "-"
    return f"{proto}:{port} <- {src}"


def collect_security_groups(ec2, region):
    rows = []
    paginator = ec2.get_paginator("describe_security_groups")
    for page in paginator.paginate():
        for sg in page["SecurityGroups"]:
            inbound = [format_rule(r) for r in sg.get("IpPermissions", [])]
            outbound = [format_rule(r) for r in sg.get("IpPermissionsEgress", [])]
            rows.append({
                "Region": region,
                "GroupId": sg["GroupId"],
                "GroupName": sg["GroupName"],
                "VpcId": sg.get("VpcId", ""),
                "Description": sg.get("Description", ""),
                "InboundRules": " | ".join(inbound) if inbound else "(none)",
                "OutboundRules": " | ".join(outbound) if outbound else "(none)",
                "Tags": tags_to_str(sg.get("Tags")),
            })
    return rows


# ---------------- VPC ----------------
def collect_vpcs(ec2, region):
    rows = []
    for vpc in ec2.describe_vpcs()["Vpcs"]:
        rows.append({
            "Region": region,
            "VpcId": vpc["VpcId"],
            "Name": name_from_tags(vpc.get("Tags")),
            "CidrBlock": vpc["CidrBlock"],
            "State": vpc["State"],
            "IsDefault": vpc["IsDefault"],
            "Tags": tags_to_str(vpc.get("Tags")),
        })
    return rows


# ---------------- Internet Gateways ----------------
def collect_igws(ec2, region):
    rows = []
    for igw in ec2.describe_internet_gateways()["InternetGateways"]:
        attachments = igw.get("Attachments", [])
        vpc_id = attachments[0]["VpcId"] if attachments else ""
        state = attachments[0]["State"] if attachments else "detached"
        rows.append({
            "Region": region,
            "InternetGatewayId": igw["InternetGatewayId"],
            "Name": name_from_tags(igw.get("Tags")),
            "AttachedVpcId": vpc_id,
            "State": state,
            "Tags": tags_to_str(igw.get("Tags")),
        })
    return rows


# ---------------- NAT Gateways ----------------
def collect_nat_gws(ec2, region):
    rows = []
    paginator = ec2.get_paginator("describe_nat_gateways")
    for page in paginator.paginate():
        for nat in page["NatGateways"]:
            public_ips = "; ".join(
                addr.get("PublicIp", "") for addr in nat.get("NatGatewayAddresses", [])
            )
            rows.append({
                "Region": region,
                "NatGatewayId": nat["NatGatewayId"],
                "Name": name_from_tags(nat.get("Tags")),
                "VpcId": nat.get("VpcId", ""),
                "SubnetId": nat.get("SubnetId", ""),
                "State": nat["State"],
                "PublicIp": public_ips,
                "Tags": tags_to_str(nat.get("Tags")),
            })
    return rows


# ---------------- Transit Gateways ----------------
def collect_tgws(ec2, region):
    rows = []
    try:
        paginator = ec2.get_paginator("describe_transit_gateways")
        for page in paginator.paginate():
            for tgw in page["TransitGateways"]:
                rows.append({
                    "Region": region,
                    "TransitGatewayId": tgw["TransitGatewayId"],
                    "Name": name_from_tags(tgw.get("Tags")),
                    "State": tgw["State"],
                    "OwnerId": tgw.get("OwnerId", ""),
                    "Description": tgw.get("Description", ""),
                    "Tags": tags_to_str(tgw.get("Tags")),
                })
    except ClientError as e:
        print(f"  ! Skipping Transit Gateways: {e}")
    return rows


# ---------------- S3 (global, not per-region) ----------------
def collect_s3(session):
    s3 = session.client("s3")
    rows = []
    buckets = s3.list_buckets()["Buckets"]
    for b in buckets:
        name = b["Name"]
        try:
            loc = s3.get_bucket_location(Bucket=name).get("LocationConstraint")
            region = loc if loc else "us-east-1"
        except ClientError:
            region = "unknown"

        try:
            tags = s3.get_bucket_tagging(Bucket=name).get("TagSet", [])
            tag_str = tags_to_str(tags)
        except ClientError:
            tag_str = ""

        try:
            lc = s3.get_bucket_lifecycle_configuration(Bucket=name).get("Rules", [])
            lc_str = "; ".join(
                f"{r.get('ID','rule')}:{r.get('Status')}" for r in lc
            )
        except ClientError:
            lc_str = "(none)"

        rows.append({
            "BucketName": name,
            "Region": region,
            "CreationDate": str(b["CreationDate"]),
            "Tags": tag_str,
            "LifecycleRules": lc_str,
        })
    return rows


def get_all_regions(session):
    ec2 = session.client("ec2", region_name="us-east-1")
    return [r["RegionName"] for r in ec2.describe_regions()["Regions"]]


def main():
    parser = argparse.ArgumentParser(description="AWS resource inventory to CSV")
    parser.add_argument("--region", default=None, help="Single region (e.g. us-east-1)")
    parser.add_argument("--all-regions", action="store_true", help="Scan all enabled regions")
    parser.add_argument("--profile", default=None, help="AWS CLI profile name")
    parser.add_argument("--outdir", default="./aws_inventory_output", help="Output directory")
    args = parser.parse_args()

    if not args.region and not args.all_regions:
        args.region = "us-east-1"
        print("No --region given, defaulting to us-east-1. Use --all-regions to scan everything.")

    os.makedirs(args.outdir, exist_ok=True)

    try:
        session = boto3.Session(profile_name=args.profile) if args.profile else boto3.Session()
    except NoCredentialsError:
        print("No AWS credentials found.")
        sys.exit(1)

    regions = get_all_regions(session) if args.all_regions else [args.region]

    all_ec2, all_sg, all_vpc, all_igw, all_nat, all_tgw = [], [], [], [], [], []

    for region in regions:
        print(f"Scanning {region} ...")
        ec2 = session.client("ec2", region_name=region)
        try:
            all_ec2 += collect_ec2(ec2, region)
            all_sg += collect_security_groups(ec2, region)
            all_vpc += collect_vpcs(ec2, region)
            all_igw += collect_igws(ec2, region)
            all_nat += collect_nat_gws(ec2, region)
            all_tgw += collect_tgws(ec2, region)
        except ClientError as e:
            print(f"  ! Error in {region}: {e}")

    print("Collecting S3 (global) ...")
    s3_rows = collect_s3(session)

    print("\nWriting per-resource CSVs:")
    write_resource_group(args.outdir, "ec2_instances", all_ec2, "InstanceId", "Name")
    write_resource_group(args.outdir, "security_groups", all_sg, "GroupId", "GroupName")
    write_resource_group(args.outdir, "vpcs", all_vpc, "VpcId", "Name")
    write_resource_group(args.outdir, "internet_gateways", all_igw, "InternetGatewayId", "Name")
    write_resource_group(args.outdir, "nat_gateways", all_nat, "NatGatewayId", "Name")
    write_resource_group(args.outdir, "transit_gateways", all_tgw, "TransitGatewayId", "Name")
    write_resource_group(args.outdir, "s3_buckets", s3_rows, "BucketName")

    print(f"\nDone. {len(regions)} region(s) scanned at {datetime.now().isoformat(timespec='seconds')}")
    print(f"Output folder: {os.path.abspath(args.outdir)}")


if __name__ == "__main__":
    main()
