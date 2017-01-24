namespace SuiteCRMClient
{
    public class CrmEntity
    {
        private readonly string _moduleName;
        private readonly string _entityId;

        public CrmEntity(string moduleName, string entityId)
        {
            _moduleName = moduleName;
            _entityId = entityId;
        }

        public string ModuleName { get { return _moduleName; } }

        public string EntityId { get { return _entityId; } }
    }
}
