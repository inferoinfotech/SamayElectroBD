// policyController.js
const Policy = require('../../models/v2/policy.model');
const ClientPolicy = require('../../models/v2/clientPolicy.model');
const logger = require('../../utils/logger');
const mongoose = require('mongoose');

// Create a new policy
exports.createPolicy = async (req, res) => {
  try {
    const { name, policies, effectiveDate, isActive } = req.body;

    if (!name || !policies || !Array.isArray(policies) || policies.length === 0) {
      return res.status(400).json({
        message: 'Policy name and at least one sub-policy are required',
      });
    }

    // Validate each sub-policy
    for (const policy of policies) {
      if (!policy.key || policy.value === undefined) {
        return res.status(400).json({
          message: 'Each sub-policy must have a key and value',
        });
      }
    }

    const newPolicy = new Policy({
      name,
      policies,
      effectiveDate: effectiveDate || new Date(),
      isActive: isActive !== undefined ? isActive : true,
    });

    await newPolicy.save();

    logger.info(`New Policy created: ${name}`);
    res.status(201).json({
      message: 'Policy created successfully',
      policy: newPolicy,
    });
  } catch (error) {
    if (error.code === 11000) {
      logger.error(`Policy name already exists: ${req.body.name}`);
      return res.status(409).json({
        message: 'Policy name already exists',
      });
    }
    logger.error(`Error creating policy: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Get all policies
exports.getAllPolicies = async (req, res) => {
  try {
    const { isActive } = req.query;
    const query = {};

    if (isActive !== undefined) {
      query.isActive = isActive === 'true';
    }

    const policies = await Policy.find(query).sort({ createdAt: -1 });

    logger.info(`Retrieved ${policies.length} policies`);
    res.status(200).json({ policies });
  } catch (error) {
    logger.error(`Error retrieving policies: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Get a single policy by ID
exports.getPolicyById = async (req, res) => {
  try {
    const { policyId } = req.params;

    if (!mongoose.Types.ObjectId.isValid(policyId)) {
      return res.status(400).json({ message: 'Invalid policy ID format' });
    }

    const policy = await Policy.findById(policyId);

    if (!policy) {
      logger.warn(`Policy not found: ${policyId}`);
      return res.status(404).json({ message: 'Policy not found' });
    }

    logger.info(`Retrieved policy: ${policyId}`);
    res.status(200).json({ policy });
  } catch (error) {
    logger.error(`Error retrieving policy: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Update a policy
exports.updatePolicy = async (req, res) => {
  try {
    const { policyId } = req.params;
    const { name, policies, effectiveDate, isActive } = req.body;

    if (!mongoose.Types.ObjectId.isValid(policyId)) {
      return res.status(400).json({ message: 'Invalid policy ID format' });
    }

    const policy = await Policy.findById(policyId);

    if (!policy) {
      logger.warn(`Policy not found: ${policyId}`);
      return res.status(404).json({ message: 'Policy not found' });
    }

    // Validate policies array if provided
    if (policies && Array.isArray(policies)) {
      for (const policyItem of policies) {
        if (!policyItem.key || policyItem.value === undefined) {
          return res.status(400).json({
            message: 'Each sub-policy must have a key and value',
          });
        }
      }
    }

    // Update fields
    if (name !== undefined) policy.name = name;
    if (policies !== undefined) policy.policies = policies;
    if (effectiveDate !== undefined) policy.effectiveDate = effectiveDate;
    if (isActive !== undefined) policy.isActive = isActive;

    await policy.save();

    logger.info(`Policy updated: ${policyId}`);
    res.status(200).json({
      message: 'Policy updated successfully',
      policy,
    });
  } catch (error) {
    if (error.code === 11000) {
      logger.error(`Policy name already exists: ${req.body.name}`);
      return res.status(409).json({
        message: 'Policy name already exists',
      });
    }
    logger.error(`Error updating policy: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Delete a policy
exports.deletePolicy = async (req, res) => {
  try {
    const { policyId } = req.params;

    if (!mongoose.Types.ObjectId.isValid(policyId)) {
      return res.status(400).json({ message: 'Invalid policy ID format' });
    }

    const policy = await Policy.findById(policyId);

    if (!policy) {
      logger.warn(`Policy not found: ${policyId}`);
      return res.status(404).json({ message: 'Policy not found' });
    }

    // Check if policy is being used by any clients
    const clientPolicies = await ClientPolicy.find({ policyId });
    if (clientPolicies.length > 0) {
      return res.status(409).json({
        message: 'Cannot delete policy. It is assigned to one or more clients.',
        assignedClients: clientPolicies.length,
      });
    }

    await Policy.findByIdAndDelete(policyId);

    logger.info(`Policy deleted: ${policyId}`);
    res.status(200).json({ message: 'Policy deleted successfully' });
  } catch (error) {
    logger.error(`Error deleting policy: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Add a sub-policy to an existing policy
exports.addSubPolicy = async (req, res) => {
  try {
    const { policyId } = req.params;
    const { key, value, description } = req.body;

    if (!mongoose.Types.ObjectId.isValid(policyId)) {
      return res.status(400).json({ message: 'Invalid policy ID format' });
    }

    if (!key || value === undefined) {
      return res.status(400).json({
        message: 'Key and value are required for sub-policy',
      });
    }

    const policy = await Policy.findById(policyId);

    if (!policy) {
      logger.warn(`Policy not found: ${policyId}`);
      return res.status(404).json({ message: 'Policy not found' });
    }

    // Check if key already exists
    const existingSubPolicy = policy.policies.find((p) => p.key === key);
    if (existingSubPolicy) {
      return res.status(409).json({
        message: 'Sub-policy with this key already exists',
      });
    }

    policy.policies.push({ key, value, description });
    await policy.save();

    logger.info(`Sub-policy added to policy: ${policyId}`);
    res.status(200).json({
      message: 'Sub-policy added successfully',
      policy,
    });
  } catch (error) {
    logger.error(`Error adding sub-policy: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Update a sub-policy
exports.updateSubPolicy = async (req, res) => {
  try {
    const { policyId, subPolicyId } = req.params;
    const { key, value, description } = req.body;

    if (!mongoose.Types.ObjectId.isValid(policyId)) {
      return res.status(400).json({ message: 'Invalid policy ID format' });
    }

    if (!mongoose.Types.ObjectId.isValid(subPolicyId)) {
      return res.status(400).json({ message: 'Invalid sub-policy ID format' });
    }

    const policy = await Policy.findById(policyId);

    if (!policy) {
      logger.warn(`Policy not found: ${policyId}`);
      return res.status(404).json({ message: 'Policy not found' });
    }

    const subPolicy = policy.policies.id(subPolicyId);
    if (!subPolicy) {
      return res.status(404).json({ message: 'Sub-policy not found' });
    }

    // Check if key is being changed and conflicts with another sub-policy
    if (key && key !== subPolicy.key) {
      const existingSubPolicy = policy.policies.find(
        (p) => p.key === key && p._id.toString() !== subPolicyId
      );
      if (existingSubPolicy) {
        return res.status(409).json({
          message: 'Sub-policy with this key already exists',
        });
      }
    }

    if (key !== undefined) subPolicy.key = key;
    if (value !== undefined) subPolicy.value = value;
    if (description !== undefined) subPolicy.description = description;

    await policy.save();

    logger.info(`Sub-policy updated: ${subPolicyId}`);
    res.status(200).json({
      message: 'Sub-policy updated successfully',
      policy,
    });
  } catch (error) {
    logger.error(`Error updating sub-policy: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

// Delete a sub-policy
exports.deleteSubPolicy = async (req, res) => {
  try {
    const { policyId, subPolicyId } = req.params;

    if (!mongoose.Types.ObjectId.isValid(policyId)) {
      return res.status(400).json({ message: 'Invalid policy ID format' });
    }

    if (!mongoose.Types.ObjectId.isValid(subPolicyId)) {
      return res.status(400).json({ message: 'Invalid sub-policy ID format' });
    }

    const policy = await Policy.findById(policyId);

    if (!policy) {
      logger.warn(`Policy not found: ${policyId}`);
      return res.status(404).json({ message: 'Policy not found' });
    }

    const subPolicy = policy.policies.id(subPolicyId);
    if (!subPolicy) {
      return res.status(404).json({ message: 'Sub-policy not found' });
    }

    // Check if sub-policy is being used by any client policies
    const clientPolicies = await ClientPolicy.find({
      policyId,
      'policies.policyItemId': subPolicyId,
    });

    if (clientPolicies.length > 0) {
      return res.status(409).json({
        message: 'Cannot delete sub-policy. It is assigned to one or more clients.',
        assignedClients: clientPolicies.length,
      });
    }

    policy.policies.pull(subPolicyId);
    await policy.save();

    logger.info(`Sub-policy deleted: ${subPolicyId}`);
    res.status(200).json({
      message: 'Sub-policy deleted successfully',
      policy,
    });
  } catch (error) {
    logger.error(`Error deleting sub-policy: ${error.message}`);
    res.status(500).json({ message: error.message });
  }
};

