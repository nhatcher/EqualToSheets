declare const _brand: unique symbol;
export type Brand<Type, Name> = Type & { [_brand]: Name };
