// This is the type for the return result of Graph API methods
export type GraphResult<Type> = {
	success: boolean;
	error?: Error;
	value?: Type;
};
