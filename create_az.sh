#!/bin/bash

# This script is to showcase my understanding of how to handle services in Azure with Bash
# We are creating multiple VMs behind a load balancer here, using redundancy 
# Azure likes to call this "Availability Zones"

# Create a resource group
az group create --name roman_az_group  --location eastus

# Create a public IP address
# This IP address will be for the load balancer
# Notice: the Standard SKU is set to standard to be able to use Availability Zones
az network public-ip create \
    --resource-group roman_az_group \
    --name azpublicip \
    --sku standard

# Create an Azure load balancer. It uses the SKU too. 
az network lb create \
    --resource-group roman_az_group \
    --name azloadbalancer \
    --public-ip-address azpublicip \
    --frontend-ip-name frontendpool \
    --backend-pool-name backendpool \
    --sku standard

# Create the first VM
az vm create \
	--resource-group roman_az_group \
	--name zonedvm1 \
	--image ubuntults \
	--size Standard_B1ms \
	--admin-username soren \
	--generate-ssh-keys \
	--zone 1

# Create a second VM
# This VM is manually defined to be created in zone 3az vm create \
	--resource-group roman_az_group \
	--name zonedvm3 \
	--image ubuntults \
	--size Standard_B1ms \
	--admin-username soren \
	--generate-ssh-keys \
	--zone 3