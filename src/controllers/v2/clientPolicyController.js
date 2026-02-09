// clientPolicyController.js
const ClientPolicy = require('../../models/v2/clientPolicy.model');
const Policy = require('../../models/v2/policy.model');
const SubClient = require('../../models/v1/subClient.model');
const logger = require('../../utils/logger');
const mongoose = require('mongoose');

// Assign a policy to a client
exports.assignPolicyToClient = async (req, res) => {
  try {
    const { subClientId, policyId, policies, effectiveDate, isActive } = req.body;

    if (!subClientId || !policyId) {
      return res.status(400).json({
        message: 'Sub client ID and policy ID are required',
      });
    }

    if (!mongoose.Types.ObjectId.isValid(subClientId)) {
      return res.status(400).json({ message: 'Invalid sub client ID format' });
    }

    if (!mongoose.Types.ObjectId.isValid(policyId)) {
      return res.status(400).json({ message: 'Invalid policy ID format' });
    }

    // Verify client exists
    const client = await SubClient.findById(subClientId);
    if (!client) {
      return res.status(404).json({ message: 'Sub client not found' });
    }

    // Verify policy exists
    const policy = await Policy.findById(policyId);
    if (!policy) {
      return res.status(404).json({ message: 'Policy not found' });
    }

    // Check if client already has any policy assigned (one policy per client rule)
    const existingClientPolicy = await ClientPolicy.findOne({
      subClientId,
    });

    // If client already has a policy, update it with the new policy
    if (existingClientPolicy) {
      // If it's the same policy, return conflict
      if (existingClientPolicy.policyId.toString() === policyId.toString()) {
        return res.status(409).json({
          message: 'This policy is already assigned to this client',
          clientPolicy: existingClientPolicy,
        });
      }
      
      // Update the existing client policy with the new policy
      existingClientPolicy.policyId = policyId;
      
      // If policies array is provided, use it; otherwise create default entries
      if (policies && Array.isArray(policies)) {
        // Validate that all policyItemIds exist in the new policy
        const policyItemIds = policy.policies.map((p) => p._id.toString());
        for (const policyItem of policies) {
          if (!policyItem.policyItemId) {
            return res.status(400).json({
              message: 'Each policy item must have a policyItemId',
            });
          }
          if (!policyItemIds.includes(policyItem.policyItemId.toString())) {
            return res.status(400).json({
              message: `Sub-policy with ID ${policyItem.policyItemId} does not exist in the policy`,
            });
          }
        }
        existingClientPolicy.policies = policies;
      } else {
        // Create default entries for all sub-policies from the new policy
        existingClientPolicy.policies = policy.policies.map((p) => ({
          policyItemId: p._id,
          apply: true,
        }));
      }
      
      existingClientPolicy.effectiveDate = effectiveDate || new Date();
      if (isActive !== undefined) {
        existingClientPolicy.isActive = isActive;
      }
      
      await existingClientPolicy.save();
      await existingClientPolicy.populate('subClientId', 'name consumerNo');
      await existingClientPolicy.populate('policyId', 'name');
      
      logger.info(`Policy updated for client: ${subClientId} -> ${policyId}`);
      return res.status(200).json({
        message: 'Client policy updated successfully (replaced with new policy)',
        clientPolicy: existingClientPolicy,
      });
    }

    // If policies array is provided, validate it
    let policiesArray = [];
    if (policies && Array.isArray(policies)) {
      // Validate that all policyItemIds exist in the policy
      const policyItemIds = policy.policies.map((p) => p._id.toString());
      for (const policyItem of policies) {
        if (!policyItem.policyItemId) {
          return res.status(400).json({
            message: 'Each policy item must have a policyItemId',
          });
        }
        if (!policyItemIds.includes(policyItem.policyItemId.toString())) {
          return res.status(400).json({
            message: `Sub-policy with ID ${policyItem.policyItemId} does not exist in the policy`,
          });
        }
      }
      policiesArray = policies;
    } else {
      // If no policies array provided, create default entries for all sub-policies
      policiesArray = policy.policies.map((p) => ({
        policyItemId: p._id,
        apply: true,
      }));
    }

    const clientPolicy = new ClientPolicy({
      subClientId,
      policyId,
      policies: policiesArray,
      effectiveDate: effectiveDate || new Date(),
      isActive: isActive !== undefined ? isActive : true,
    });

    await clientPolicy.save();

    // Populate references for response
    await clientPolicy.populate('subClientId', 'name consumerNo');
    await clientPolicy.populate('policyId', 'name');

    logger.info(`Policy assigned to client: ${subClientId} -> ${policyId}`);
    res.status(201).json({
      message: 'Policy assigned to client successfully',
      clientPolicy,
    });
  } catch (error) {
    if (error.code === 11000) {
      logger.error(`Policy already assigned to client`);
      return res.status(409).json({
        message: 'Policy is already assigned to this client',
      });
    }
    logger.error(`Error assigning policy to client: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Get all client-policy mappings
exports.getAllClientPolicies = async (req, res) => {
  try {
    const { subClientId, policyId, isActive } = req.query;
    const query = {};

    if (subClientId) {
      if (!mongoose.Types.ObjectId.isValid(subClientId)) {
        return res.status(400).json({ message: 'Invalid sub client ID format' });
      }
      query.subClientId = subClientId;
    }

    if (policyId) {
      if (!mongoose.Types.ObjectId.isValid(policyId)) {
        return res.status(400).json({ message: 'Invalid policy ID format' });
      }
      query.policyId = policyId;
    }

    if (isActive !== undefined) {
      query.isActive = isActive === 'true';
    }

    const clientPolicies = await ClientPolicy.find(query)
      .populate('subClientId', 'name consumerNo')
      .populate('policyId', 'name policies')
      .sort({ createdAt: -1 });

    logger.info(`Retrieved ${clientPolicies.length} client-policy mappings`);
    res.status(200).json({ clientPolicies });
  } catch (error) {
    logger.error(`Error retrieving client-policy mappings: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Get a single client-policy mapping
exports.getClientPolicyById = async (req, res) => {
  try {
    const { clientPolicyId } = req.params;

    if (!mongoose.Types.ObjectId.isValid(clientPolicyId)) {
      return res.status(400).json({ message: 'Invalid client-policy ID format' });
    }

    const clientPolicy = await ClientPolicy.findById(clientPolicyId)
      .populate('subClientId', 'name consumerNo')
      .populate('policyId', 'name policies');

    if (!clientPolicy) {
      logger.warn(`Client-policy mapping not found: ${clientPolicyId}`);
      return res.status(404).json({ message: 'Client-policy mapping not found' });
    }

    logger.info(`Retrieved client-policy mapping: ${clientPolicyId}`);
    res.status(200).json({ clientPolicy });
  } catch (error) {
    logger.error(`Error retrieving client-policy mapping: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Get policies for a specific client
exports.getClientPolicies = async (req, res) => {
  try {
    const { clientId } = req.params;

    if (!mongoose.Types.ObjectId.isValid(clientId)) {
      return res.status(400).json({ message: 'Invalid client ID format' });
    }

    const clientPolicies = await ClientPolicy.find({
      subClientId: clientId,
      isActive: true,
    })
      .populate('policyId', 'name policies effectiveDate')
      .sort({ effectiveDate: -1 });

    logger.info(`Retrieved policies for client: ${clientId}`);
    res.status(200).json({ clientPolicies });
  } catch (error) {
    logger.error(`Error retrieving client policies: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Update client-policy mapping
exports.updateClientPolicy = async (req, res) => {
  try {
    const { clientPolicyId } = req.params;
    const { policies, effectiveDate, isActive } = req.body;

    if (!mongoose.Types.ObjectId.isValid(clientPolicyId)) {
      return res.status(400).json({ message: 'Invalid client-policy ID format' });
    }

    const clientPolicy = await ClientPolicy.findById(clientPolicyId).populate(
      'policyId'
    );

    if (!clientPolicy) {
      logger.warn(`Client-policy mapping not found: ${clientPolicyId}`);
      return res.status(404).json({ message: 'Client-policy mapping not found' });
    }

    // If policies array is provided, validate it
    if (policies && Array.isArray(policies)) {
      const policy = await Policy.findById(clientPolicy.policyId);
      const policyItemIds = policy.policies.map((p) => p._id.toString());

      for (const policyItem of policies) {
        if (!policyItem.policyItemId) {
          return res.status(400).json({
            message: 'Each policy item must have a policyItemId',
          });
        }
        if (!policyItemIds.includes(policyItem.policyItemId.toString())) {
          return res.status(400).json({
            message: `Sub-policy with ID ${policyItem.policyItemId} does not exist in the policy`,
          });
        }
      }
      clientPolicy.policies = policies;
    }

    if (effectiveDate !== undefined) clientPolicy.effectiveDate = effectiveDate;
    if (isActive !== undefined) clientPolicy.isActive = isActive;

    await clientPolicy.save();

    await clientPolicy.populate('subClientId', 'name consumerNo');
    await clientPolicy.populate('policyId', 'name policies');

    logger.info(`Client-policy mapping updated: ${clientPolicyId}`);
    res.status(200).json({
      message: 'Client-policy mapping updated successfully',
      clientPolicy,
    });
  } catch (error) {
    logger.error(`Error updating client-policy mapping: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Update a specific sub-policy apply status for a client
exports.updateClientSubPolicy = async (req, res) => {
  try {
    const { clientPolicyId, subPolicyId } = req.params;
    const { apply, customValue } = req.body;

    if (!mongoose.Types.ObjectId.isValid(clientPolicyId)) {
      return res.status(400).json({ message: 'Invalid client-policy ID format' });
    }

    if (!mongoose.Types.ObjectId.isValid(subPolicyId)) {
      return res.status(400).json({ message: 'Invalid sub-policy ID format' });
    }

    const clientPolicy = await ClientPolicy.findById(clientPolicyId);

    if (!clientPolicy) {
      logger.warn(`Client-policy mapping not found: ${clientPolicyId}`);
      return res.status(404).json({ message: 'Client-policy mapping not found' });
    }

    const subPolicy = clientPolicy.policies.id(subPolicyId);
    if (!subPolicy) {
      return res.status(404).json({ message: 'Sub-policy not found in client mapping' });
    }

    if (apply !== undefined) subPolicy.apply = apply;
    if (customValue !== undefined) subPolicy.customValue = customValue;

    await clientPolicy.save();

    await clientPolicy.populate('subClientId', 'name consumerNo');
    await clientPolicy.populate('policyId', 'name policies');

    logger.info(`Client sub-policy updated: ${subPolicyId}`);
    res.status(200).json({
      message: 'Client sub-policy updated successfully',
      clientPolicy,
    });
  } catch (error) {
    logger.error(`Error updating client sub-policy: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Remove policy assignment from a client
exports.removeClientPolicy = async (req, res) => {
  try {
    const { clientPolicyId } = req.params;

    if (!mongoose.Types.ObjectId.isValid(clientPolicyId)) {
      return res.status(400).json({ message: 'Invalid client-policy ID format' });
    }

    const clientPolicy = await ClientPolicy.findById(clientPolicyId);

    if (!clientPolicy) {
      logger.warn(`Client-policy mapping not found: ${clientPolicyId}`);
      return res.status(404).json({ message: 'Client-policy mapping not found' });
    }

    await ClientPolicy.findByIdAndDelete(clientPolicyId);

    logger.info(`Client-policy mapping deleted: ${clientPolicyId}`);
    res.status(200).json({ message: 'Policy assignment removed from client successfully' });
  } catch (error) {
    logger.error(`Error removing client-policy mapping: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

